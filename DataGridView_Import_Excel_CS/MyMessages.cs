using System;
using System.Drawing;
using System.Windows.Forms;

namespace Productivity
{
    public partial class MyMessages : Form
    {
        public MyMessages(string name, string message, byte n)
        {
            InitializeComponent();
            this.Text = name;
            this.label1.Text = message;
            if (n == 1)
            {
                this.pictureBox1.Image = Productivity.Properties.Resources.OK;
            }
            else
            {
                this.pictureBox1.BackgroundImage = Productivity.Properties.Resources.Excel;
            }

            label1.MaximumSize = new Size(300, 0);
            label1.AutoSize = true;
            pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
        }

        private bool okButton = false;

        public bool OKButtonClicked
        {
            get { return okButton; }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            okButton = true;
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            okButton = false;
            this.Close();

        }

        private void MyMessages_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
