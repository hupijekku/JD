﻿using System;
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

        DataTable dt;
        string[,] taulukko;

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
                colCount = range.Columns.Count + 2;

                dgv.ColumnCount = colCount;
                try
                {
                    for (int i = 0; i < rowCount; i++)
                    {
                        DataGridViewRow row = new DataGridViewRow();
                        row.CreateCells(dgv);

                        row.Cells[0].Value = i + 1;
                        row.Cells[1].Value = "Ei paikalla";
                        row.DefaultCellStyle.BackColor = Color.Red;
                        row.DefaultCellStyle.Font = new Font("Arial", 15.0f, GraphicsUnit.Pixel);

                        for (int j = 2; j < colCount; j++)
                        {
                            if((range.Cells[i + 1, j - 1] as Excel.Range).Value2 != null)
                            {
                                row.Cells[j].Value = (range.Cells[i + 1, j - 1] as Excel.Range).Value2.ToString();
                            }
                            else
                            {
                                row.Cells[j].Value = "";
                            }
                        }

                        dgv.Rows.Add(row);
                    }
                    DGVTaulukkoon();
                }
                catch(Exception exc)
                {
                    MessageBox.Show("Jotain meni pieleen. \n" +
                        "Suosittelen poistamaan Excel-tiedostosta tyhjät rivit ja sarakkeet \n\n" +
                        "Jos ei auta, ota yhteyttä Eemeliin, ja näytä tämä: \n" + exc.ToString());
                }

                dgv.Columns[0].Width = 40;

                //string str = (string)(range.Cells[i, j] as Excel.Range).Value2;


                xlWorkbook.Close();
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
            }
            
        }

        private void muokkaus_Click(object sender, EventArgs e)
        {
            if(dgv.ReadOnly)
            {
                dgv.ReadOnly = false;
                muokkaus.ForeColor = Color.Red;
                muokkaus.Text = "Muokkaus pois";
            }
            else
            {
                dgv.ReadOnly = true;
                muokkaus.ForeColor = Color.Black;
                muokkaus.Text = "Muokkaus päälle";

            }
        }

        private void Lue_Click(object sender, EventArgs e)
        {
            Lue lue = new Lue();
            lue.mainDGV = dgv;
            lue.mainForm = this;
            lue.Show();
            //MessageBox.Show(EtsiRivi(1).ToString());
        }

        public int EtsiRivi(int ID)
        {
            //MessageBox.Show(dgv.Rows[0].Cells[0].ToString());
            for (int i = 0; i < dgv.RowCount; i++)
            {
                if(dgv.Rows[i].Cells[0].Value.ToString() == ID.ToString())
                {
                    return i;
                }
            }

            return 9999;
        }

        private void Tallenna_Click(object sender, EventArgs e)
        {
            SortDGV();
            dt = new DataTable { TableName = "Henkilöt" };
            foreach(DataGridViewColumn dgc in dgv.Columns)
            {
                dt.Columns.Add(dgc.Name);
            }
            foreach(DataGridViewRow dgr in dgv.Rows)
            {
                DataRow dRow = dt.NewRow();
                foreach(DataGridViewCell cell in dgr.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dRow);
            }
            string currentPath = System.Environment.CurrentDirectory;
            dt.WriteXml(currentPath + "\\henkilot.xml", XmlWriteMode.WriteSchema);
        }

        private void Avaa_Click(object sender, EventArgs e)
        {
            dt = new DataTable();
            string currentPath = System.Environment.CurrentDirectory;
            dt.ReadXml(currentPath + "\\henkilot.xml");

            dgv.Rows.Clear();
            dgv.Refresh();

            dgv.ColumnCount = dt.Columns.Count;
            int colCount = dgv.ColumnCount;
            int rowCount = dt.Rows.Count;

            for (int i = 0; i < rowCount; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgv);
                row.Cells[0].Value = i + 1;
                row.DefaultCellStyle.BackColor = Color.Red;
                row.DefaultCellStyle.Font = new Font("Arial", 15.0f, GraphicsUnit.Pixel);
                for (int j = 1; j < colCount; j++)
                {
                    row.Cells[j].Value = dt.Rows[i][j];
                }
                if(row.Cells[1].Value.ToString() == "Paikalla")
                {
                    row.DefaultCellStyle.BackColor = Color.Green;
                }

                dgv.Rows.Add(row);
            }
            DGVTaulukkoon();
        }

        private void LisaaRivi_Click(object sender, EventArgs e)
        {
            try
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgv);
                row.Cells[0].Value = dgv.Rows.Count + 1;
                row.Cells[1].Value = "Ei Paikalla";
                row.DefaultCellStyle.BackColor = Color.Red;
                row.DefaultCellStyle.Font = new Font("Arial", 15.0f, GraphicsUnit.Pixel);
                dgv.Rows.Add(row);
            }
            catch(Exception exc)
            {
                MessageBox.Show("Ei taulukkoa mihin lisätä riviä? \n\ntai näytä Eemelille: \n\n" + exc.ToString()); 
            }

        }

        public void Save()
        {
            Tallenna.PerformClick();
        }

        public void SortDGV()
        {
            dgv.Sort(dgv.Columns[0], ListSortDirection.Ascending);
        }

        public void DGVTaulukkoon()
        {
            SortDGV();
            int ColCount = dgv.ColumnCount - 1;
            int RowCount = dgv.RowCount;
            taulukko = new string[ColCount, RowCount];
            for (int i = 0; i < ColCount; i++)
            {
                for (int j = 0; j < RowCount; j++)
                {
                    taulukko[i, j] = dgv.Rows[j].Cells[i + 1].Value.ToString();
                }
            }
        }

        public void TaulukkoDGVeen()
        {
            dgv.Rows.Clear();
            dgv.Refresh();
            int colCount = taulukko.GetLength(0);
            int rowCount = taulukko.GetLength(1);

            for (int i = 0; i < rowCount; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgv);
                row.Cells[0].Value = i + 1;
                row.DefaultCellStyle.BackColor = Color.Red;
                row.DefaultCellStyle.Font = new Font("Arial", 15.0f, GraphicsUnit.Pixel);
                for (int j = 0; j < colCount; j++)
                {
                    row.Cells[j + 1].Value = taulukko[j, i];
                }
                if (row.Cells[1].Value.ToString() == "Paikalla")
                {
                    row.DefaultCellStyle.BackColor = Color.Green;
                }

                dgv.Rows.Add(row);
            }
            DGVTaulukkoon();
        }

        private void Haku_TextChanged(object sender, EventArgs e)
        {
            try
            {
                TaulukkoDGVeen();
                string hae = Haku.Text;
                bool found = false;
                for (int i = 0; i < dgv.RowCount; i++)
                {
                    found = false;
                    for (int j = 0; j < dgv.ColumnCount; j++)
                    {
                        if (dgv.Rows[i].Cells[j].Value.ToString().ToLower().Contains(hae.ToLower()))
                        {
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        dgv.Rows.RemoveAt(i);
                        i--;
                    }
                }
            }
            catch
            {
                Haku.Text = "";
            }
        }
    }
}
