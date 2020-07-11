namespace Design_Dashboard_Modern
{
    partial class fmbaocao
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea4 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend4 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series4 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea3 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend3 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.panel1 = new System.Windows.Forms.Panel();
            this.charttheongaythang = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.panel3 = new System.Windows.Forms.Panel();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.panel5 = new System.Windows.Forms.Panel();
            this.DoanhThuTheoNgay = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.chart2 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.ThongKePhieuTra = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.charttheongaythang)).BeginInit();
            this.panel3.SuspendLayout();
            this.panel5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DoanhThuTheoNgay)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ThongKePhieuTra)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.charttheongaythang);
            this.panel1.Location = new System.Drawing.Point(29, 39);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(596, 464);
            this.panel1.TabIndex = 0;
            // 
            // charttheongaythang
            // 
            chartArea1.Name = "ChartArea1";
            this.charttheongaythang.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.charttheongaythang.Legends.Add(legend1);
            this.charttheongaythang.Location = new System.Drawing.Point(0, 0);
            this.charttheongaythang.Name = "charttheongaythang";
            series1.BackImageTransparentColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            series1.BorderColor = System.Drawing.Color.Yellow;
            series1.ChartArea = "ChartArea1";
            series1.Color = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            series1.Legend = "Legend1";
            series1.Name = "Doanh Thu Theo Thang";
            this.charttheongaythang.Series.Add(series1);
            this.charttheongaythang.Size = new System.Drawing.Size(596, 464);
            this.charttheongaythang.TabIndex = 0;
            this.charttheongaythang.Text = "chart1";
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.dateTimePicker2);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.dateTimePicker1);
            this.panel3.Controls.Add(this.label1);
            this.panel3.Location = new System.Drawing.Point(29, 3);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1394, 30);
            this.panel3.TabIndex = 3;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(1012, 3);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(248, 22);
            this.dateTimePicker2.TabIndex = 3;
            this.dateTimePicker2.ValueChanged += new System.EventHandler(this.dateTimePicker2_ValueChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(687, 3);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(319, 22);
            this.label2.TabIndex = 2;
            this.label2.Text = "Chọn Thời Gian Thống kê Của Năm :\r\n";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(433, 3);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(248, 22);
            this.dateTimePicker1.TabIndex = 1;
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(3, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(424, 22);
            this.label1.TabIndex = 0;
            this.label1.Text = "Chọn Thời Gian Thống kê Theo Tháng Của Năm :\r\n";
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.DoanhThuTheoNgay);
            this.panel5.Location = new System.Drawing.Point(29, 509);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(596, 406);
            this.panel5.TabIndex = 5;
            // 
            // DoanhThuTheoNgay
            // 
            chartArea2.Name = "ChartArea1";
            this.DoanhThuTheoNgay.ChartAreas.Add(chartArea2);
            legend2.Name = "Legend1";
            this.DoanhThuTheoNgay.Legends.Add(legend2);
            this.DoanhThuTheoNgay.Location = new System.Drawing.Point(3, 28);
            this.DoanhThuTheoNgay.Name = "DoanhThuTheoNgay";
            this.DoanhThuTheoNgay.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Chocolate;
            series2.ChartArea = "ChartArea1";
            series2.Legend = "Legend1";
            series2.Name = "Doanh Thu Theo Nam";
            this.DoanhThuTheoNgay.Series.Add(series2);
            this.DoanhThuTheoNgay.Size = new System.Drawing.Size(593, 367);
            this.DoanhThuTheoNgay.TabIndex = 0;
            this.DoanhThuTheoNgay.Text = "chart1";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.ThongKePhieuTra);
            this.panel2.Location = new System.Drawing.Point(686, 39);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(737, 464);
            this.panel2.TabIndex = 6;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.chart2);
            this.panel4.Location = new System.Drawing.Point(695, 523);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(724, 354);
            this.panel4.TabIndex = 7;
            // 
            // chart2
            // 
            chartArea4.Name = "ChartArea1";
            this.chart2.ChartAreas.Add(chartArea4);
            legend4.Name = "Legend1";
            this.chart2.Legends.Add(legend4);
            this.chart2.Location = new System.Drawing.Point(3, 3);
            this.chart2.Name = "chart2";
            series4.ChartArea = "ChartArea1";
            series4.Legend = "Legend1";
            series4.Name = "Trả Nhà Cung Cấp Theo Năm";
            this.chart2.Series.Add(series4);
            this.chart2.Size = new System.Drawing.Size(722, 348);
            this.chart2.TabIndex = 0;
            this.chart2.Text = "chart2";
            // 
            // ThongKePhieuTra
            // 
            chartArea3.Name = "ChartArea1";
            this.ThongKePhieuTra.ChartAreas.Add(chartArea3);
            legend3.Name = "Legend1";
            this.ThongKePhieuTra.Legends.Add(legend3);
            this.ThongKePhieuTra.Location = new System.Drawing.Point(9, 3);
            this.ThongKePhieuTra.Name = "ThongKePhieuTra";
            this.ThongKePhieuTra.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Berry;
            series3.BackImageTransparentColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            series3.BorderColor = System.Drawing.Color.Yellow;
            series3.ChartArea = "ChartArea1";
            series3.Color = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            series3.Legend = "Legend1";
            series3.Name = "Tiền trả Khi Lấy Hàng Theo Tháng";
            this.ThongKePhieuTra.Series.Add(series3);
            this.ThongKePhieuTra.Size = new System.Drawing.Size(724, 458);
            this.ThongKePhieuTra.TabIndex = 1;
            this.ThongKePhieuTra.Text = "chart1";
            // 
            // fmbaocao
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(30)))), ((int)(((byte)(30)))), ((int)(((byte)(46)))));
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel5);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel1);
            this.Name = "fmbaocao";
            this.Size = new System.Drawing.Size(1467, 918);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.charttheongaythang)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.DoanhThuTheoNgay)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ThongKePhieuTra)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.DataVisualization.Charting.Chart charttheongaythang;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.DataVisualization.Charting.Chart DoanhThuTheoNgay;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart2;
        private System.Windows.Forms.DataVisualization.Charting.Chart ThongKePhieuTra;
    }
}
