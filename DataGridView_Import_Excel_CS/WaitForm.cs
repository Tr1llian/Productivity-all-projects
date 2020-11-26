using System.Windows.Forms;

namespace Productivity
{
    public partial class WaitForm : Form
    {
        public int ProgressBarValue
        {
            get { return (this.progressBar1.Value); }
            set { if (value == 100) this.Close();
                else
                    this.progressBar1.Value = value; }
        }

        public WaitForm()
        {
            InitializeComponent();
            progressBar1.Value = 0;
          
        }


     
    }
}
