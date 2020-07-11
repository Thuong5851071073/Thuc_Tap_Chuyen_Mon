namespace Design_Dashboard_Modern
{
    partial class XemHoadon
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.datahoadon = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tk = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.datahoadon)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // datahoadon
            // 
            this.datahoadon.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.datahoadon.Location = new System.Drawing.Point(254, 198);
            this.datahoadon.Name = "datahoadon";
            this.datahoadon.RowTemplate.Height = 24;
            this.datahoadon.Size = new System.Drawing.Size(755, 292);
            this.datahoadon.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.tk);
            this.panel1.Controls.Add(this.textBox1);
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Location = new System.Drawing.Point(152, 30);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1028, 80);
            this.panel1.TabIndex = 1;
            // 
            // tk
            // 
            this.tk.AutoSize = true;
            this.tk.Font = new System.Drawing.Font("Times New Roman", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tk.ForeColor = System.Drawing.Color.White;
            this.tk.Location = new System.Drawing.Point(242, 46);
            this.tk.Name = "tk";
            this.tk.Size = new System.Drawing.Size(194, 19);
            this.tk.TabIndex = 3;
            this.tk.Text = "Tìm kiếm theo tình trạng";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(451, 11);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(215, 54);
            this.textBox1.TabIndex = 2;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Yellow;
            this.panel2.Location = new System.Drawing.Point(242, 71);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(424, 1);
            this.panel2.TabIndex = 1;
            // 
            // XemHoadon
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(30)))), ((int)(((byte)(46)))));
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.datahoadon);
            this.Name = "XemHoadon";
            this.Size = new System.Drawing.Size(1467, 918);
            this.Load += new System.EventHandler(this.XemHoadon_Load);
            ((System.ComponentModel.ISupportInitialize)(this.datahoadon)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView datahoadon;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label tk;
    }
}
