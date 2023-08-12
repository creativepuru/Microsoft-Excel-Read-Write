using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadWriteExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonRead_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            string currentDirectory = Directory.GetCurrentDirectory();
            string filePath = Path.Combine(currentDirectory, "TempReader.xlsx");

            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Assuming data is on the first sheet

            int rowCount = worksheet.UsedRange.Rows.Count;
            List<string> cellValues = new List<string>();

            // Read data from all cells in the first column
            for (int row = 1; row <= rowCount; row++)
            {
                Excel.Range cell = worksheet.Cells[row, 1];
                string cellValue = cell.Value != null ? cell.Value.ToString() : "";
                cellValues.Add(cellValue);
            }

            // Define a column for the DataGridView
            dataGridView.Columns.Clear();
            dataGridView.Columns.Add("CellValue", "Cell Value");

            // Display all data in DataGridView
            dataGridView.Rows.Clear();
            dataGridView.Rows.AddRange(cellValues.Select(value => new DataGridViewRow { Cells = { new DataGridViewTextBoxCell { Value = value } } }).ToArray());

            // Close workbook and Excel application
            workbook.Close(false);
            excelApp.Quit();

        }

        private void buttonWrite_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            string currentDirectory = Directory.GetCurrentDirectory();
            string filePath = Path.Combine(currentDirectory, "TempReader.xlsx");

            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Accessing the first sheet

            worksheet.Cells[5, 1] = "E";

            // Save the workbook
            workbook.Save();

            // Close workbook and Excel application
            workbook.Close(false);
            excelApp.Quit();

            MessageBox.Show("Data written to Excel successfully.");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
    }
}
