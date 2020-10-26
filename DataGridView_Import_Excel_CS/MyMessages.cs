using System;
using System.Windows.Forms;

namespace Productivity
{
    public partial class MyMessages : Form
    {
        public MyMessages()
        {
            InitializeComponent();
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
    }
}
