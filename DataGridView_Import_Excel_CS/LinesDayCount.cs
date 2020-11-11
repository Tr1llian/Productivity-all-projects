using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;

namespace Productivity
{
    public partial class LinesDayCount : Form
    {
        public LinesDayCount(List<Saloon> cars)
        {
            InitializeComponent();
            flowLayoutPanel1.FlowDirection = FlowDirection.TopDown;
            flowLayoutPanel1.AutoScroll = true;
            flowLayoutPanel1.WrapContents = false;
            foreach (Saloon car in cars)
            {
                Label name = new Label();
                name.Text = car.ProjectName;
                //name.Location = new Point(10, 10);
                
                Label LinesCount = new Label();
                LinesCount.Text = "Кількість бригад";
                LinesCount.Location = new Point(10, 40);
                
                Label DaysDes = new Label();
                DaysDes.Text ="Кількість днів";
                DaysDes.Size = new Size(100, 50);
                DaysDes.Location = new Point(10, 60);
                
                TextBox lines = new TextBox();
                lines.Size = new Size(50, 30);
                lines.Text = car.lines.ToString();
                lines.Location = new Point(200, 40);
               
                TextBox days = new TextBox();
                days.Text = car.days.ToString();
                days.Size = new Size(50, 30);
                days.Location = new Point(200, 60);
                
                GroupBox g = new GroupBox();
                g.Size = new Size(290, 100);
                g.Controls.Add(name);
                g.Controls.Add(lines);
                g.Controls.Add(days);
                g.Controls.Add(DaysDes);
                g.Controls.Add(LinesCount);
                
                
                flowLayoutPanel1.Controls.Add(g);
            }
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void LinesDayCount_Load(object sender, System.EventArgs e)
        {

        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }
}
