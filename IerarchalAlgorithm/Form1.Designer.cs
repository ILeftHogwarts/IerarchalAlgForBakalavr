namespace IerarchalAlgorithm
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.UpIerarch = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.AlgorithmCalcButton = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.UpIerarch)).BeginInit();
            this.SuspendLayout();
            // 
            // UpIerarch
            // 
            chartArea2.Name = "ChartArea1";
            this.UpIerarch.ChartAreas.Add(chartArea2);
            legend2.Name = "Legend1";
            this.UpIerarch.Legends.Add(legend2);
            this.UpIerarch.Location = new System.Drawing.Point(12, 12);
            this.UpIerarch.Name = "UpIerarch";
            this.UpIerarch.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.EarthTones;
            series2.ChartArea = "ChartArea1";
            series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            series2.Legend = "Legend1";
            series2.Name = "Series1";
            this.UpIerarch.Series.Add(series2);
            this.UpIerarch.Size = new System.Drawing.Size(647, 425);
            this.UpIerarch.TabIndex = 0;
            this.UpIerarch.Text = "UpIerarch";
            this.UpIerarch.Click += new System.EventHandler(this.UpIerarch_Click);
            // 
            // AlgorithmCalcButton
            // 
            this.AlgorithmCalcButton.Location = new System.Drawing.Point(750, 12);
            this.AlgorithmCalcButton.Name = "AlgorithmCalcButton";
            this.AlgorithmCalcButton.Size = new System.Drawing.Size(149, 63);
            this.AlgorithmCalcButton.TabIndex = 1;
            this.AlgorithmCalcButton.Text = "Вверх";
            this.AlgorithmCalcButton.UseVisualStyleBackColor = true;
            this.AlgorithmCalcButton.Click += new System.EventHandler(this.AlgorithmCalcButton_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(751, 82);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(148, 64);
            this.button1.TabIndex = 2;
            this.button1.TabStop = false;
            this.button1.Text = "Вниз";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(750, 153);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(149, 65);
            this.button2.TabIndex = 3;
            this.button2.Text = "Двойная";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(911, 462);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.AlgorithmCalcButton);
            this.Controls.Add(this.UpIerarch);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.UpIerarch)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart UpIerarch;
        private System.Windows.Forms.Button AlgorithmCalcButton;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}

