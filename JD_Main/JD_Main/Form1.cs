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
using OfficeOpenXml;

namespace JD_Main
{
    public partial class Form1 : Form
    {
        ExcelPackage exl;
        
        //Datatable to store DataGridView in to easily write it to the .xml file
        DataTable dt;
        //Keep the DGV in memory for search function
        string[,] taulukko;

        public Form1()
        {
            InitializeComponent();
        }

        private void tuoExcel_click(object sender, EventArgs e)
        {
            try
            {
                var filePath = string.Empty;
                int rowCount = 0;
                int colCount = 0;

                //Open the excel file and write it's contents to the DGV
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel file (*.xls*)|*.xls*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    exl = new ExcelPackage(new System.IO.FileInfo(ofd.FileName));

                    //First worksheet
                    ExcelWorksheet worksheet = exl.Workbook.Worksheets[1];

                    rowCount = worksheet.Dimension.End.Row;
                    colCount = worksheet.Dimension.End.Column + 2;

                    dgv.ColumnCount = colCount;
                    try
                    {
                        for (int i = 0; i < rowCount; i++)
                        {
                            DataGridViewRow row = new DataGridViewRow();
                            row.CreateCells(dgv);

                            //Manually adding the first 2 columns
                            //Incrementing by 1, because ID-cards start from 001 -->
                            row.Cells[0].Value = i + 1;
                            row.Cells[1].Value = "Ei paikalla";
                            row.DefaultCellStyle.BackColor = Color.Red;
                            row.DefaultCellStyle.Font = new Font("Arial", 15.0f, GraphicsUnit.Pixel);

                            //Creating columns 2 -->
                            for (int j = 2; j < colCount; j++)
                            {
                                //i + 1, because Excel starts from 1
                                //j + 1 - the 2 first columns => j - 1
                                //Replace null with "" to prevent crashing
                                if ((worksheet.Cells[i + 1, j - 1]).Value?.ToString().Trim() != null)
                                {
                                    row.Cells[j].Value = (worksheet.Cells[i + 1, j - 1]).Value?.ToString().Trim();
                                }
                                else
                                {
                                    row.Cells[j].Value = "";
                                }
                            }

                            dgv.Rows.Add(row);
                        }
                        //Write the data to array to save it for search-function
                        DGVTaulukkoon();
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show("Jotain meni pieleen. \n" +
                            "Suosittelen poistamaan Excel-tiedostosta tyhjät rivit ja sarakkeet \n\n" +
                            "Jos ei auta, ota yhteyttä Eemeliin, ja näytä tämä: \n" + exc.ToString());
                    }

                    dgv.Columns[0].Width = 40;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Virheilmoitus: \n\n" + ex.ToString());
            }
            
            
        }

        private void muokkaus_Click(object sender, EventArgs e)
        {
            //Edit mode => Allow editing cell values.
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
            //Open ID-reading form.
            Lue lue = new Lue();
            lue.mainDGV = dgv;
            lue.mainForm = this;
            lue.Show();
        }

        public int EtsiRivi(int ID)
        {
            //Find row index that matches the ID (ID is unique)
            for (int i = 0; i < dgv.RowCount; i++)
            {
                if(dgv.Rows[i].Cells[0].Value.ToString() == ID.ToString())
                {
                    return i;
                }
            }

            //if (rowNum != 9999) (Lue.cs)
            return 9999;
        }

        private void Tallenna_Click(object sender, EventArgs e)
        {
            //Sorting before saving to easily keep first column as integers
            //Reading from .xml would change them to strings, which ruins sorting (e.g. 1, 10, 2, 3)
            SortDGV();

            //Convert DGV to DataTable for dt.WriteXml()
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
            //currentDirectory\henkilot.xml
            dt.WriteXml(currentPath + "\\henkilot.xml", XmlWriteMode.WriteSchema);
        }

        private void Avaa_Click(object sender, EventArgs e)
        {
            //Read data to DataTable
            dt = new DataTable();
            string currentPath = System.Environment.CurrentDirectory;
            dt.ReadXml(currentPath + "\\henkilot.xml");

            //Clear old data to prevent duplicates
            dgv.Rows.Clear();
            dgv.Refresh();


            dgv.ColumnCount = dt.Columns.Count;
            int colCount = dgv.ColumnCount;
            int rowCount = dt.Rows.Count;

            //Fill DGV with data from DataTable
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
            //Update Array values for search-function.
            DGVTaulukkoon();
        }

        private void LisaaRivi_Click(object sender, EventArgs e)
        {
            try
            {
                //Add a new row to the end of the DGV, match style and fill first 2 columns.
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
            //Executes Tallenna_Click(); Why did I make this a separate function?
            Tallenna.PerformClick();
        }

        public void SortDGV()
        {
            dgv.Sort(dgv.Columns[0], ListSortDirection.Ascending);
        }

        public void DGVTaulukkoon()
        {
            //Sort before saving to the array. Easier to keep IDs as integers
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
            //Clear old data from DGV
            dgv.Rows.Clear();
            dgv.Refresh();
            int colCount = taulukko.GetLength(0);
            int rowCount = taulukko.GetLength(1);

            //Loop array to fill DGV, add first column manually to keep IDs as integers.
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
            //Save just in case, not sure if necessary. Doesn't affect performance much.
            DGVTaulukkoon();
        }

        private void Haku_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //Reset DGV to original
                TaulukkoDGVeen();

                string hae = Haku.Text;
                bool found = false;

                //Loop DGV, remove row if it doesn't contain search value.
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
                //This runs twice because it changes the text which triggers the event...
                //meh..
                Haku.Text = "";
            }
        }
    }
}
