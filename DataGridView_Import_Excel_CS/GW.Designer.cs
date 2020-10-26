namespace DataGridView_Import_Excel
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.MailBoxList = new System.Windows.Forms.ListBox();
            this.AddMail = new System.Windows.Forms.Button();
            this.DeleteMail = new System.Windows.Forms.Button();
            this.MailTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SentMail = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // MailBoxList
            // 
            this.MailBoxList.FormattingEnabled = true;
            this.MailBoxList.ItemHeight = 16;
            this.MailBoxList.Location = new System.Drawing.Point(94, 27);
            this.MailBoxList.Name = "MailBoxList";
            this.MailBoxList.Size = new System.Drawing.Size(500, 308);
            this.MailBoxList.TabIndex = 6;
            // 
            // AddMail
            // 
            this.AddMail.Location = new System.Drawing.Point(94, 379);
            this.AddMail.Name = "AddMail";
            this.AddMail.Size = new System.Drawing.Size(498, 46);
            this.AddMail.TabIndex = 4;
            this.AddMail.Text = "Додати";
            this.AddMail.UseVisualStyleBackColor = true;
            this.AddMail.Click += new System.EventHandler(this.AddMail_Click);
            // 
            // DeleteMail
            // 
            this.DeleteMail.Location = new System.Drawing.Point(94, 431);
            this.DeleteMail.Name = "DeleteMail";
            this.DeleteMail.Size = new System.Drawing.Size(500, 46);
            this.DeleteMail.TabIndex = 5;
            this.DeleteMail.Text = "Видалити";
            this.DeleteMail.UseVisualStyleBackColor = true;
            this.DeleteMail.Click += new System.EventHandler(this.Button1_Click);
            // 
            // MailTextBox
            // 
            this.MailTextBox.Location = new System.Drawing.Point(178, 345);
            this.MailTextBox.Name = "MailTextBox";
            this.MailTextBox.Size = new System.Drawing.Size(414, 22);
            this.MailTextBox.TabIndex = 3;
            this.MailTextBox.TextChanged += new System.EventHandler(this.MailTextBox_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(89, 341);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 25);
            this.label2.TabIndex = 8;
            this.label2.Text = "E-mail";
            // 
            // SentMail
            // 
            this.SentMail.Location = new System.Drawing.Point(92, 483);
            this.SentMail.Name = "SentMail";
            this.SentMail.Size = new System.Drawing.Size(500, 46);
            this.SentMail.TabIndex = 2;
            this.SentMail.Text = "Надіслати";
            this.SentMail.UseVisualStyleBackColor = true;
            this.SentMail.Click += new System.EventHandler(this.SentMail_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::Productivity.Properties.Resources.Bader_feel;
            this.ClientSize = new System.Drawing.Size(660, 589);
            this.Controls.Add(this.SentMail);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.MailTextBox);
            this.Controls.Add(this.DeleteMail);
            this.Controls.Add(this.AddMail);
            this.Controls.Add(this.MailBoxList);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Надіслати звіт";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ListBox MailBoxList;
        private System.Windows.Forms.Button AddMail;
        private System.Windows.Forms.Button DeleteMail;
        private System.Windows.Forms.TextBox MailTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button SentMail;
    }
}