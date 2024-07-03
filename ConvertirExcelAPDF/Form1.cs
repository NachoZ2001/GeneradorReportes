using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;

namespace ConvertirExcelAPDF
{
    public partial class Form1 : Form
    {
        private List<string> excelFiles;

        public Form1()
        {
            InitializeComponent();
            excelFiles = new List<string>();

            pictureBox.Visible = false;
        }


        private void btnSelectFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel Files|*.xlsx;*.xls"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                excelFiles.AddRange(openFileDialog.FileNames);
                foreach (var file in openFileDialog.FileNames)
                {
                    listBoxFiles.Items.Add(file);
                }
            }
        }

        private async void btnConvertToPdf_Click(object sender, EventArgs e)
        {
            pictureBox.Visible = true;

            foreach (var excelFile in excelFiles)
            {
                await Task.Run(() => ConvertExcelToPdf(excelFile));
            }

            MessageBox.Show("Conversion complete!");

            pictureBox.Visible = false;
        }

        private void ConvertExcelToPdf(string excelFilePath)
        {
            // Load the Excel file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(excelFilePath);

            // Set PDF save options
            workbook.ConverterSetting.SheetFitToPage = true;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.Orientation = PageOrientationType.Landscape;
            }

            // Define the output PDF file path
            string pdfFilePath = Path.ChangeExtension(excelFilePath, ".pdf");

            // Save the Excel file as PDF
            workbook.SaveToFile(pdfFilePath, Spire.Xls.FileFormat.PDF);
        }
    }
}
