using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Data.SqlClient;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;





namespace Proyecto_1_final_Mistral
{
    public partial class Form1 : Form
    {
        string conexion = "Server=LAPTOP-R27RH76P\\SQLEXPRESS01;Database=Para_Proyecto_1;Trusted_Connection=True;";

        public object Office { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)

        {
            string prompt = textBox1.Text;
            if (string.IsNullOrWhiteSpace(prompt))
            {
                MessageBox.Show("Por favor, escribe un tema.");
                return;
            }

            string apiKey = "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"; // 👈 Pega aquí tu API Key de OpenRouter
            string url = "https://openrouter.ai/api/v1/chat/completions";

            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
            httpClient.DefaultRequestHeaders.Add("HTTP-Referer", "https://tupagina.com"); // requerido por OpenRouter
            httpClient.DefaultRequestHeaders.Add("X-Title", "MiAppIA"); // nombre de tu app

            var requestBody = new
            {
                model = "mistralai/mistral-7b-instruct:free",
                messages = new[]
                {
            new { role = "user", content = prompt }
        }
            };

            var content = new StringContent(JsonConvert.SerializeObject(requestBody), Encoding.UTF8, "application/json");

            try
            {
                var response = await httpClient.PostAsync(url, content);
                string responseBody = await response.Content.ReadAsStringAsync();

                dynamic result = JsonConvert.DeserializeObject(responseBody);
                string respuesta = result?.choices?[0]?.message?.content;

                if (string.IsNullOrEmpty(respuesta))
                {
                    MessageBox.Show("No se recibió una respuesta válida.");
                    return;
                }

                richTextBox1.Text = respuesta.Trim();
                using (SqlConnection conn = new SqlConnection(conexion))
                {
                    conn.Open();
                    string insertQuery = "INSERT INTO Investigaciones (Prompt, Respuesta) VALUES (@prompt, @respuesta)";
                    using (SqlCommand cmd = new SqlCommand(insertQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@prompt", prompt);
                        cmd.Parameters.AddWithValue("@respuesta", respuesta);
                        cmd.ExecuteNonQuery();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al conectarse con OpenRouter:\n" + ex.Message);
            }
        }
        private void GenerarDocumentoWord()
        {
            string connectionString = @"Server=LAPTOP-R27RH76P\SQLEXPRESS01;Database=Para_Proyecto_1;Trusted_Connection=True;"; // Cambia aquí tu cadena de conexión
            string query = "SELECT TOP 1 * FROM Investigaciones ORDER BY id DESC"; // Cambia 'ResultadosIA' por el nombre real de tu tabla

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataReader reader = command.ExecuteReader();

                if (reader.Read())
                {
                    string promt = reader["Prompt"].ToString();
                    string respuestas = reader["Respuesta"].ToString();
                    string fecha = reader["Fecha"].ToString();

                    var wordApp = new Word.Application();
                    var documento = wordApp.Documents.Add();
                    var parrafo = documento.Content.Paragraphs.Add();

                    parrafo.Range.Text = "Reporte generado por IA";
                    parrafo.Range.InsertParagraphAfter();

                    parrafo.Range.Text = $"Prompt: {promt}";
                    parrafo.Range.InsertParagraphAfter();

                    parrafo.Range.Text = $"Respuesta generada:\n{respuestas}";
                    parrafo.Range.InsertParagraphAfter();

                    parrafo.Range.Text = $"Fecha: {fecha}";
                    parrafo.Range.InsertParagraphAfter();

                    string carpeta = @"C:\Users\Public\ReportesIA";
                    if (!Directory.Exists(carpeta))
                    {
                        Directory.CreateDirectory(carpeta);
                    }

                    string ruta = Path.Combine(carpeta, "ReporteIA.docx"); // Cambia esta ruta si quieres
                    documento.SaveAs2(ruta);
                    documento.Close();
                    wordApp.Quit();

                    System.Diagnostics.Process.Start("explorer.exe", carpeta);


                    System.Diagnostics.Process.Start("winword.exe", ruta);


                    MessageBox.Show("¡Documento Word creado con éxito en:\n" + ruta + "!");
                }
                else
                {
                    MessageBox.Show("No se encontraron datos en la base.");
                }
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {
            GenerarDocumentoWord();
        }

        private void GenerarPresentacionPowerPoint()
        {
            string connectionString = @"Server=LAPTOP-R27RH76P\SQLEXPRESS01;Database=Para_Proyecto_1;Trusted_Connection=True;";
            string query = "SELECT TOP 1 * FROM Investigaciones ORDER BY Id DESC";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand command = new SqlCommand(query, connection);
                SqlDataReader reader = command.ExecuteReader();

                if (reader.Read())
                {
                    string promt = reader["Prompt"].ToString();
                    string respuesta = reader["Respuesta"].ToString();
                    string fecha = reader["Fecha"].ToString();

                    var pptApp = new PowerPoint.Application();
                    var presentacion = pptApp.Presentations.Add();

                    var slide1 = presentacion.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);
                    slide1.Shapes.Title.TextFrame.TextRange.Text = "Reporte generado por IA";
                    slide1.Shapes.Placeholders[2].TextFrame.TextRange.Text = "Fecha: " + fecha;

                    var slide2 = presentacion.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutText);
                    slide2.Shapes.Placeholders[1].TextFrame.TextRange.Text = "Prompt";
                    slide2.Shapes.Placeholders[2].TextFrame.TextRange.Text = promt;

                    var slide3 = presentacion.Slides.Add(3, PowerPoint.PpSlideLayout.ppLayoutText);
                    slide3.Shapes.Placeholders[1].TextFrame.TextRange.Text = "Respuesta generada por IA";
                    slide3.Shapes.Placeholders[2].TextFrame.TextRange.Text = respuesta;

                    string carpeta = @"C:\Users\Public\ReportesIA";
                    if (!Directory.Exists(carpeta))
                    {
                        Directory.CreateDirectory(carpeta);
                    }

                    string ruta = @"C:\Users\Public\ReporteIA.pptx";
                    presentacion.SaveAs(ruta);
                    presentacion.Close();
                    pptApp.Quit();

                    System.Diagnostics.Process.Start("explorer.exe", carpeta);

                    MessageBox.Show("¡Presentación PowerPoint generada con éxito en:\n" + ruta + "!");
                }
                else
                {
                    MessageBox.Show("No se encontraron datos en la base.");
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            GenerarPresentacionPowerPoint();
        }
    }
}

