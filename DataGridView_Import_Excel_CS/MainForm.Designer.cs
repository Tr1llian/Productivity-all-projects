namespace DataGridView_Import_Excel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnSelect = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.btnPrint = new System.Windows.Forms.Button();
            this.printDocument1 = new System.Drawing.Printing.PrintDocument();
            this.button2 = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(16, 71);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.Size = new System.Drawing.Size(1280, 482);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView1_CellContentClick);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.OpenFileDialog1_FileOk);
            // 
            // btnSelect
            // 
            this.btnSelect.Location = new System.Drawing.Point(16, 15);
            this.btnSelect.Margin = new System.Windows.Forms.Padding(4);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(131, 49);
            this.btnSelect.TabIndex = 2;
            this.btnSelect.Text = "Вибрати файл";
            this.btnSelect.UseVisualStyleBackColor = true;
            this.btnSelect.Click += new System.EventHandler(this.BtnSelect_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(155, 15);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(133, 49);
            this.button1.TabIndex = 3;
            this.button1.Text = "експорт в  excel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.Location = new System.Drawing.Point(296, 15);
            this.btnPrint.Margin = new System.Windows.Forms.Padding(4);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(133, 47);
            this.btnPrint.TabIndex = 4;
            this.btnPrint.Text = "друк";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.BtnPrint_Click_1);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(437, 15);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(133, 47);
            this.button2.TabIndex = 5;
            this.button2.Text = "Надіслати звіт";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(1312, 569);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnSelect);
            this.Controls.Add(this.dataGridView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MinimumSize = new System.Drawing.Size(1327, 605);
            this.Name = "Form1";
            this.Text = "Продуктивність";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnSelect;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnPrint;
        private System.Drawing.Printing.PrintDocument printDocument1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}

