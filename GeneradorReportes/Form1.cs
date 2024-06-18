using System;
using System.Collections.Generic;
using System.Drawing; 
using System.Globalization;
using System.IO;
using System.Linq;
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

        }

        private void SeleccionarArchivo(TextBox textBox)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Archivos Excel|*.xlsx;";
                openFileDialog.Title = "Seleccionar el archivo Excel";

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

                var diccionario = ArmarDiccionario(textBoxExcel.Text, int.Parse(textBoxPeriodo.Text));
                diccionario = TransformarProyectos(diccionario, int.Parse(textBoxPeriodo.Text));

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Archivos Excel|*.xlsx";
                saveFileDialog.Title = "Guardar Informe como";
                saveFileDialog.FileName = $"List. Venc. {textBoxNombre.Text}";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string rutaArchivo = saveFileDialog.FileName;
                    CrearInformeExcel(diccionario, rutaArchivo, textBoxNombre.Text);

                    MessageBox.Show("Proceso completado");
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
                if (string.IsNullOrEmpty(rutaArchivo) || !rutaArchivo.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException("La ruta del archivo debe tener una extensión .xlsx");
                }

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Informe");

                    var imagePath = "C:\\Users\\w1zar\\Downloads\\imagen.png";
                    if (File.Exists(imagePath))
                    {
                        var image = worksheet.AddPicture(imagePath)
                            .MoveTo(worksheet.Cell("B1"), new Point(10, 10))
                            .Scale(0.5);
                    }

                    worksheet.Cell(4, 1).Value = nombre;
                    worksheet.Cell(4, 1).Style.Font.Bold = true;
                    worksheet.Cell(4, 1).Style.Font.FontSize = 14;
                    worksheet.Cell(4, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;

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

                    worksheet.Columns().AdjustToContents();

                    int fila = 9;
                    int filaInicio = 9;
                    foreach (var proyecto in proyectos)
                    {
                        if (proyecto.Value == null || proyecto.Value.Count == 0)
                        {
                            continue;
                        }

                        foreach (var (asunto, periodo, vencimiento) in proyecto.Value)
                        {
                            worksheet.Cell(fila, 1).Value = proyecto.Key;
                            worksheet.Cell(fila, 2).Value = asunto;
                            worksheet.Cell(fila, 3).Value = periodo;
                            worksheet.Cell(fila, 4).Value = vencimiento;
                            fila++;
                        }

                        var range = worksheet.Range($"A{filaInicio}:A{fila - 1}");
                        range.Merge();
                        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                        worksheet.Cell(fila, 1).Value = " ";
                        worksheet.Cell(fila, 2).Value = " ";
                        worksheet.Cell(fila, 3).Value = " ";
                        worksheet.Cell(fila, 4).Value = " ";

                        var pieRange = worksheet.Range($"A{fila}:D{fila}");
                        pieRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#941BB5");
                        pieRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        pieRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        pieRange.Style.Font.FontColor = XLColor.White;
                        pieRange.Style.Font.Bold = true;

                        fila++;
                        filaInicio = fila;
                    }

                    var usedRange = worksheet.Range(worksheet.Cell(8, 1), worksheet.LastCellUsed());
                    foreach (var cell in usedRange.Cells())
                    {
                        cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                        cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                        cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    }

                    worksheet.Column(1).Width = PixelsToCharacterWidth(308);
                    worksheet.Column(2).Width = PixelsToCharacterWidth(275);
                    worksheet.Column(3).Width = PixelsToCharacterWidth(110);
                    worksheet.Column(4).Width = PixelsToCharacterWidth(120);

                    workbook.SaveAs(rutaArchivo);
                }

                Console.WriteLine("Informe creado exitosamente en: " + rutaArchivo);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocurrió un error al crear el informe de Excel: " + ex.Message);
            }
        }


        static double PixelsToCharacterWidth(int pixels)
        {
            return (pixels / 7.0) * 0.75;
        }
    }
}
