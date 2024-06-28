using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace GeneradorReportes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            pictureBoxRuedaCargando.Visible = false;
        }

        private void SeleccionarArchivo(TextBox textBox)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Archivos Excel|*.xlsx;*.csv";
                openFileDialog.Title = "Seleccionar el archivo Excel o CSV";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox.Text = openFileDialog.FileName;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SeleccionarArchivo(textBoxExcel);
        }

        private void buttonProcesar_Click(object sender, EventArgs e)
        {
            pictureBoxRuedaCargando.Visible = true;

            string rutaArchivo = textBoxExcel.Text;
            if (Path.GetExtension(rutaArchivo).ToLower() == ".csv")
            {
                rutaArchivo = ConvertirCsvAXlsx(rutaArchivo);
            }

            var diccionario = ArmarDiccionario(rutaArchivo, int.Parse(textBoxPeriodo.Text));
            diccionario = TransformarProyectos(diccionario, int.Parse(textBoxPeriodo.Text));

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Archivos Excel|*.xlsx";
            saveFileDialog.Title = "Guardar Informe como";
            saveFileDialog.FileName = $"List. Venc. {textBoxNombre.Text}";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string rutaInforme = saveFileDialog.FileName;
                CrearInformeExcel(diccionario, rutaInforme, textBoxNombre.Text);

                MessageBox.Show("Proceso completado");

                pictureBoxRuedaCargando.Visible = false;
            }
        }

        static string ObtenerNombreProyecto(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return input;
            }

            char[] delimiters = { '-', '/' };
            int index = input.IndexOfAny(delimiters);

            return index == -1 ? input : input.Substring(0, index);
        }

        static string ProcesarAsunto(string input)
        {
            if (input == "Sicore Pago a cuenta")
            {
                return "Pago a cta. Sicore";
            }
            else if (input == "Sicore DD.JJ")
            {
                return "Sicore";
            }
            else if (input == "Sipres - 1ºQ")
            {
                return "Sipres - 1º Quincena";
            }
            else if (input == "Sipres - 2ºQ")
            {
                return "Sipres - 2º Quincena";
            }
            else if (input == "Autonomos")
            {
                return "Autónomos";
            }
            else
            {
                return input;
            }
        }

        static Dictionary<string, List<(string, string, string)>> ArmarDiccionario(string rutaExcel, int periodoObjetivo)
        {
            var diccionario = new Dictionary<string, List<(string, string, string)>>();

            using (var workbook = new XLWorkbook(rutaExcel))
            {
                var worksheet = workbook.Worksheet(1);
                int ultimaFila = worksheet.LastRowUsed().RowNumber();

                for (int fila = 2; fila <= ultimaFila; fila++)
                {
                    string nombreProyecto = ObtenerNombreProyecto(worksheet.Cell(fila, 2).GetString());
                    string asunto = ProcesarAsunto(worksheet.Cell(fila, 4).GetString());
                    string periodo = worksheet.Cell(fila, 5).Value.ToString();
                    string vencimiento = worksheet.Cell(fila, 6).GetString();

                    if (ObtenerMes(vencimiento) != periodoObjetivo)
                    {
                        vencimiento = worksheet.Cell(fila, 7).GetString();
                    }

                    if (!diccionario.ContainsKey(nombreProyecto))
                    {
                        diccionario[nombreProyecto] = new List<(string, string, string)>();
                    }

                    diccionario[nombreProyecto].Add((asunto, periodo, vencimiento));
                }
            }

            return diccionario;
        }

        static Dictionary<string, List<(string asunto, string periodo, string vencimiento)>> TransformarProyectos(Dictionary<string, List<(string asunto, string periodo, string vencimiento)>> proyectos, int periodoObjetivo)
        {
            var transformado = new Dictionary<string, List<(string asunto, string periodo, string vencimiento)>>();

            foreach (var proyecto in proyectos)
            {
                var listaTransformada = proyecto.Value
                    .Where(x => ObtenerMes(x.vencimiento) == periodoObjetivo)
                    .OrderBy(x => DateTime.Parse(x.vencimiento))
                    .Select(x => (
                        x.asunto,
                        TransformarPeriodo(x.periodo),
                        TransformarVencimiento(x.vencimiento)
                    )).ToList();

                transformado.Add(proyecto.Key, listaTransformada);
            }

            return transformado;
        }

        static int ObtenerMes(string periodo)
        {
            if (DateTime.TryParse(periodo, out DateTime fecha))
            {
                return fecha.Month;
            }
            return 0;
        }

        static string TransformarPeriodo(string periodo)
        {
            if (DateTime.TryParse(periodo, out DateTime fecha))
            {
                return fecha.ToString("MMM-yy", CultureInfo.InvariantCulture).ToLower();
            }
            return string.Empty;
        }

        static string TransformarVencimiento(string vencimiento)
        {
            if (DateTime.TryParse(vencimiento, out DateTime fecha))
            {
                return fecha.ToString("dd-MMM", CultureInfo.InvariantCulture).ToLower();
            }
            return string.Empty;
        }

        static void CrearInformeExcel(Dictionary<string, List<(string asunto, string periodo, string vencimiento)>> proyectos, string rutaArchivo, string nombre)
        {
            try
            {
                // Verificación de la ruta del archivo
                if (string.IsNullOrEmpty(rutaArchivo) || !rutaArchivo.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException("La ruta del archivo debe tener una extensión .xlsx");
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Informe");

                    // Ruta de la imagen desde los recursos
                    var imageResource = Properties.Resources.imagen; // Asumiendo que la imagen está en los recursos
                    if (imageResource != null)
                    {
                        using (var stream = new MemoryStream())
                        {
                            imageResource.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                            stream.Seek(0, SeekOrigin.Begin);
                            var image = worksheet.AddPicture(stream, "imagen.png")
                                .MoveTo(worksheet.Cell("B1"), new Point(10, 10))
                                .Scale(0.5);
                        }
                    }

                    // Añadir nombre
                    worksheet.Cell(4, 1).Value = nombre;
                    worksheet.Cell(4, 1).Style.Font.Bold = true;
                    worksheet.Cell(4, 1).Style.Font.FontSize = 14;
                    worksheet.Cell(4, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    worksheet.Cell(4, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Encabezados de columnas
                    worksheet.Cell(8, 1).Value = "";
                    worksheet.Cell(8, 2).Value = "Asunto";
                    worksheet.Cell(8, 3).Value = "Periodo";
                    worksheet.Cell(8, 4).Value = "Vto. Pago";

                    var headerRange = worksheet.Range("A8:D8");
                    headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#941BB5");
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    headerRange.Style.Font.FontColor = XLColor.White;
                    headerRange.Style.Font.Bold = true;

                    // Ajustar anchos de columnas
                    worksheet.Column(1).Width = 280 / 7.5;
                    worksheet.Column(2).Width = 250 / 7.5;
                    worksheet.Column(3).Width = 80 / 7.5;
                    worksheet.Column(4).Width = 80 / 7.5;

                    int fila = 9;
                    int filaInicial = 9;
                    int filaFinal = filaInicial;

                    // Añadir datos del proyecto
                    foreach (var proyecto in proyectos)
                    {
                        if (proyecto.Value == null || proyecto.Value.Count == 0)
                        {
                            continue;
                        }

                        int filaInicioProyecto = fila; // Guardar el inicio del grupo del proyecto

                        foreach (var (asunto, periodo, vencimiento) in proyecto.Value)
                        {
                            worksheet.Cell(fila, 1).Value = proyecto.Key;
                            worksheet.Cell(fila, 2).Value = asunto;
                            worksheet.Cell(fila, 3).Value = periodo;
                            worksheet.Cell(fila, 4).Value = vencimiento;
                            fila++;
                        }

                        // Combinar celdas para el nombre del proyecto
                        var projectRange = worksheet.Range($"A{filaInicioProyecto}:A{fila - 1}");
                        projectRange.Merge();
                        projectRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        projectRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                        // Actualizar filaFinal después de agregar los datos del proyecto
                        filaFinal = fila - 1;

                        // Añadir fila de separación
                        worksheet.Cell(fila, 1).Value = " ";
                        worksheet.Cell(fila, 2).Value = " ";
                        worksheet.Cell(fila, 3).Value = " ";
                        worksheet.Cell(fila, 4).Value = " ";

                        var pieRange = worksheet.Range($"A{fila}:D{fila}");
                        pieRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#941BB5");
                        pieRange.Style.Font.FontColor = XLColor.White;
                        pieRange.Style.Font.Bold = true;

                        fila++;
                        filaInicial = fila;
                    }

                    // Aplicar bordes interiores a todo el rango de datos al final del foreach
                    if (filaFinal >= 9) // Verifica que se hayan añadido datos
                    {
                        var dataRange = worksheet.Range($"A8:D{filaFinal + 1}");
                        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; // Bordes exteriores finos
                        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;  // Bordes interiores finos
                    }

                    workbook.SaveAs(rutaArchivo);
                }
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("Error en la ruta del archivo: " + ex.Message);
            }
            catch (IOException ex)
            {
                MessageBox.Show("Error de E/S: " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al crear el archivo Excel: " + ex.Message);
            }
        }

        static string ConvertirCsvAXlsx(string rutaCsv)
        {
            string rutaXlsx = Path.ChangeExtension(rutaCsv, ".xlsx");

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Datos");

                using (var reader = new StreamReader(rutaCsv, Encoding.UTF8))
                {
                    int fila = 1;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        var values = line.Split(';');

                        for (int columna = 0; columna < values.Length; columna++)
                        {
                            worksheet.Cell(fila, columna + 1).Value = values[columna];
                        }

                        fila++;
                    }
                }

                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(rutaXlsx);
            }

            return rutaXlsx;
        }
    }
}
