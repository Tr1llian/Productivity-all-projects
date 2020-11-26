using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System;

namespace Productivity
{
    public partial class LinesDayCount : Form
    {
        readonly List<Saloon> CarsCopy;
        public LinesDayCount(ref List<Saloon> cars)
        {
            CarsCopy = cars;
            InitializeComponent();
            flowLayoutPanel1.FlowDirection = FlowDirection.TopDown;
            flowLayoutPanel1.AutoScroll = true;
            flowLayoutPanel1.WrapContents = false;
            foreach (Saloon car in cars)
            {
                Label name = new Label
                {
                    Text = car.ProjectName
                };
                //name.Location = new Point(10, 10);

                Label LinesCount = new Label
                {
                    Text = "Кількість бригад",
                    Location = new Point(10, 40)
                };

                Label DaysDes = new Label
                {
                    Text = "Кількість днів",
                    Size = new Size(100, 50),
                    Location = new Point(10, 60)
                };

                TextBox lines = new TextBox
                {
                    Name = car.ProjectName.ToString(),
                    Size = new Size(50, 30),
                    Text = car.lines.ToString(),
                    Location = new Point(200, 40)
                };

                TextBox days = new TextBox
                {
                    Name = "Days " + car.ProjectName.ToString(),
                    Text = car.days.ToString(),
                    Size = new Size(50, 30),
                    Location = new Point(200, 60)
                };

                GroupBox g = new GroupBox
                {
                    Size = new Size(290, 100)
                };

                g.Controls.Add(name);
                g.Controls.Add(lines);
                g.Controls.Add(days);
                g.Controls.Add(DaysDes);
                g.Controls.Add(LinesCount);


                flowLayoutPanel1.Controls.Add(g);
            }
        }

        private void FlowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void LinesDayCount_Load(object sender, System.EventArgs e)
        {

        }

        private void Button1_Click(object sender, System.EventArgs e)
        {
            foreach(Saloon a in CarsCopy)
            {
                foreach (GroupBox c in flowLayoutPanel1.Controls)
                {

                    foreach (Control b in c.Controls)
                    {
                        if(b.Name == a.ProjectName)
                        {
                            if(a.lines !=Convert.ToInt32(b.Text.ToString()))
                            {
                                a.UpdateLines(Convert.ToInt32(b.Text.ToString()));
                            }
                        }
                        if(b.Name.Contains("Days")&& b.Name.Contains(a.ProjectName))
                        {
                            if( a.days != Convert.ToInt32(b.Text.ToString()))
                            {
                                a.UpdateDays(Convert.ToInt32(b.Text.ToString()));
                            }
                        }
                       
                        //Console.WriteLine(b.ToString());
                        //Console.WriteLine(b.Text);
                    }
                }
                
            }

            Close();
        }
    }
}
