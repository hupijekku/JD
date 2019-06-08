namespace JD_Main
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuTop = new System.Windows.Forms.MenuStrip();
            this.tiedostoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.Tallenna = new System.Windows.Forms.ToolStripMenuItem();
            this.Avaa = new System.Windows.Forms.ToolStripMenuItem();
            this.tuoExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.muokkaus = new System.Windows.Forms.ToolStripMenuItem();
            this.Lue = new System.Windows.Forms.ToolStripMenuItem();
            this.LisaaRivi = new System.Windows.Forms.ToolStripMenuItem();
            this.menuTop.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // menuTop
            // 
            this.menuTop.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tiedostoToolStripMenuItem,
            this.muokkaus,
            this.Lue,
            this.LisaaRivi});
            this.menuTop.Location = new System.Drawing.Point(0, 0);
            this.menuTop.Name = "menuTop";
            this.menuTop.Size = new System.Drawing.Size(615, 24);
            this.menuTop.TabIndex = 1;
            this.menuTop.Text = "Menu";
            // 
            // tiedostoToolStripMenuItem
            // 
            this.tiedostoToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.Tallenna,
            this.Avaa,
            this.tuoExcel});
            this.tiedostoToolStripMenuItem.Name = "tiedostoToolStripMenuItem";
            this.tiedostoToolStripMenuItem.Size = new System.Drawing.Size(65, 20);
            this.tiedostoToolStripMenuItem.Text = "Tiedosto";
            // 
            // Tallenna
            // 
            this.Tallenna.Name = "Tallenna";
            this.Tallenna.Size = new System.Drawing.Size(180, 22);
            this.Tallenna.Text = "Tallenna";
            this.Tallenna.Click += new System.EventHandler(this.Tallenna_Click);
            // 
            // Avaa
            // 
            this.Avaa.Name = "Avaa";
            this.Avaa.Size = new System.Drawing.Size(180, 22);
            this.Avaa.Text = "Avaa";
            this.Avaa.Click += new System.EventHandler(this.Avaa_Click);
            // 
            // tuoExcel
            // 
            this.tuoExcel.Name = "tuoExcel";
            this.tuoExcel.Size = new System.Drawing.Size(180, 22);
            this.tuoExcel.Text = "Tuo Excelistä";
            this.tuoExcel.Click += new System.EventHandler(this.tuoExcel_click);
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.AllowUserToOrderColumns = true;
            this.dgv.AllowUserToResizeRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgv.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dgv.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.Raised;
            this.dgv.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable;
            this.dgv.ColumnHeadersHeight = 20;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgv.EnableHeadersVisualStyles = false;
            this.dgv.Location = new System.Drawing.Point(12, 27);
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.RowHeadersVisible = false;
            this.dgv.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dgv.Size = new System.Drawing.Size(591, 351);
            this.dgv.TabIndex = 2;
            // 
            // muokkaus
            // 
            this.muokkaus.Name = "muokkaus";
            this.muokkaus.Size = new System.Drawing.Size(135, 20);
            this.muokkaus.Text = "Muokkaus Päälle/Pois";
            this.muokkaus.Click += new System.EventHandler(this.muokkaus_Click);
            // 
            // Lue
            // 
            this.Lue.Name = "Lue";
            this.Lue.Size = new System.Drawing.Size(38, 20);
            this.Lue.Text = "Lue";
            this.Lue.Click += new System.EventHandler(this.Lue_Click);
            // 
            // LisaaRivi
            // 
            this.LisaaRivi.Name = "LisaaRivi";
            this.LisaaRivi.Size = new System.Drawing.Size(64, 20);
            this.LisaaRivi.Text = "Lisää rivi";
            this.LisaaRivi.Click += new System.EventHandler(this.LisaaRivi_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(615, 390);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.menuTop);
            this.MainMenuStrip = this.menuTop;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Main";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.menuTop.ResumeLayout(false);
            this.menuTop.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuTop;
        private System.Windows.Forms.ToolStripMenuItem tiedostoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem Tallenna;
        private System.Windows.Forms.ToolStripMenuItem Avaa;
        private System.Windows.Forms.ToolStripMenuItem tuoExcel;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.ToolStripMenuItem muokkaus;
        private System.Windows.Forms.ToolStripMenuItem Lue;
        private System.Windows.Forms.ToolStripMenuItem LisaaRivi;
    }
}

