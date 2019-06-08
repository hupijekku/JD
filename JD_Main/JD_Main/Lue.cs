using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JD_Main
{
    public partial class Lue : Form
    {
        public Lue()
        {
            InitializeComponent();
        }

        //Get main form and DGV
        //Used functions from mainForm: EtsiRivi(), Save().
        public DataGridView mainDGV { get; set; }
        public Form1 mainForm { get; set; }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //Find row with the ID being read and change it's status
                lb_error.Text = "";
                int rowNum = mainForm.EtsiRivi(Int32.Parse(textBox1.Text));
                if (rowNum != 9999)
                {
                    DataGridViewRow row = mainDGV.Rows[rowNum];
                    if(row.Cells[1].Value.ToString() == "Paikalla")
                    {
                        row.DefaultCellStyle.BackColor = Color.Red;
                        row.Cells[1].Value = "Ei Paikalla";
                    } else
                    {
                        row.DefaultCellStyle.BackColor = Color.Green;
                        row.Cells[1].Value = "Paikalla";
                    }
                    textBox1.Text = "";
                    //Save changes.
                    mainForm.Save();
                }
                //ID too large or 9999, yeah the program doesn't work if you have 10 000 people
                //That seems very likely..
                else lb_error.Text = "ID:tä ei löydy";
            }
            catch
            {
                textBox1.Text = "";
                lb_error.Text = "Virheellinen syöte.";
            }
        }
    }
}
