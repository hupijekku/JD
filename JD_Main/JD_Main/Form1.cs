using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace JD_Main
{
    public partial class Form1 : Form
    {

        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel.Worksheet xlWorksheet;
        Excel.Range range;

        public Form1()
        {
            InitializeComponent();
        }

        private void tuoExcel_click(object sender, EventArgs e)
        {
            var filePath = string.Empty;
            int rowCount = 0;
            int colCount = 0;

            xlApp = new Excel.Application();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel file (*.xls*)|*.xls*";
            if(ofd.ShowDialog() == DialogResult.OK)
            {
                filePath = ofd.FileName;
                xlWorkbook = xlApp.Workbooks.Open(
                    filePath, 0, true, 5, "", "", true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                    "\t", false, false, 0, true, 1, 0
                );

                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);

                range = xlWorksheet.UsedRange;
                rowCount = range.Rows.Count;
                colCount = range.Columns.Count;

                dgv.ColumnCount = colCount;

                for (int i = 0; i < rowCount; i++)
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dgv);

                    for (int j = 0; j < colCount; j++)
                    {
                        row.Cells[j].Value = (string)(range.Cells[i + 1, j + 1] as Excel.Range).Value2;
                    }

                    dgv.Rows.Add(row);
                }


                //string str = (string)(range.Cells[i, j] as Excel.Range).Value2;


                xlWorkbook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
            }
            
        }
    }
}
