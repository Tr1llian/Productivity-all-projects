using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace Productivity
{
    class SaloonBMWhiga : Saloon
    {
        public SaloonBMWhiga(Saloon G11, string name)
        {
            RBcount = G11.RBcount;
            RBtime = G11.RBtime;
            RCtime = G11.RC40time + G11.RC100time;
            RC40count = G11.RC40count;
            RC100count = G11.RC100count;
            RCcount = G11.RC100count + G11.RC40count;
            ProjectName = name;
            InitLines();
        }

        public override double AvgTime()
        {
            double AllPcs = RC100count * 2 + RC40count + RBcount;
            double AllPcsTime = RCtime + RBtime;
            if(AllPcs !=0)
            {
                return AllPcsTime / AllPcs;
            }
            else
            {
                return 0;
            }
            //return Math.Round(((double)(RCtime + RBtime) / (double)(RC100count * 2 + RC40count + RBcount)), 3);
        }

        public override void CreateRow(ref DataRow row1)
        {
            row1["Проект"] = "BMW higa";
            row1["Кількість чохлів"] = "\n RB = " + RBcount
                + "\n" + " RC40 = " + RC40count
                + "\n" + " RC100 = " + RC100count +
                "\n" + "Загальна кількість=" + GeneralCount() + "\n";
            row1["Загальний час"] = "\n RB time = " + RBtime
                + "\n" + " RC time = " + RCtime +
                "\n" + "Загальна час=" + GeneralTime() + "\n";
            row1["Час на одну штуку"] = "\n RB time for pcs= " + Math.Round(PartTime(RBtime, RBcount), 3)
                + "\n" + " RC time for pcs= " + Math.Round(PartTime(RCtime, RC100count * 2 + RC40count), 3) + "\n";
            row1["Час на салон"] = Math.Round(((PartTime(RBtime, RBcount) * 2) + PartTime(RCtime, RCcount)) / 0.35);
            row1["Кількість салонів"] = Math.Floor((RBcount + RC100count * 2 + RC40count) / Coef);
            row1["Середній час на одну штуку"] = Math.Round(((double)(RCtime + RBtime) / (double)(RC100count * 2 + RC40count + RBcount)), 3);
            row1["Коефіцієнт/кількість компонентів"] = Coef;
            row1["Кількість компонент помножено на середній на одну штуку"] = Math.Round(Coef * AvgTime(), 3);
            row1["Prod. sets planned"] = Math.Round(480 / (Coef * AvgTime()), 3);
            row1["Кількість бригад"] = lines;
            row1["Кількість днів"] = days;
            row1["Кількість бригад soll"] = lines * days;
            row1["Кількість бригад ist"] = UniqueLines();
            row1["Коефіцієнт"] = Math.Round(PartTime(UniqueLines(), lines * days), 3);
            row1["дні"] = Math.Round(PartTime(UniqueLines(), lines * days) * days, 3);
        }

        public override double GeneralCount()
        {
            return FCcount + FBcount + RC100count * 2 + RC40count + RBcount;
        }

        public override double GeneralTime()
        {
            return FCtime + FBtime + RCtime + RBtime;
        }

        public override void ParseExcel(DataRow row)
        {
            //Console.WriteLine("Not needed to Parse :D");
            if (Convert.ToInt32(row[3].ToString()) >= 5000000)
            {
                LineDay l = new LineDay(Convert.ToInt32(row[2].ToString()), row[0].ToString());
                if (!LD.Contains(l))
                {
                    LD.Add(l);
                }
            }
        }

        public override void InitLines()
        {
            string fileName = Path.Combine(Application.StartupPath, "Settings.xml");
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(fileName);
            XmlElement xRoot = xDoc.DocumentElement;
            foreach (XmlNode xnode in xRoot)
            {
                // отримуємо атрибут name
                if (xnode.Attributes.Count > 0)
                {
                    XmlNode attr = xnode.Attributes.GetNamedItem("name");
                    if (attr != null)
                        Console.WriteLine(attr.Value);
                }
                // обходимо всі дочірні елементи user
                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    // Якщо вузол - company
                    if (childnode.Name == "BMWhiga")
                    {
                        lines =Convert.ToInt32( childnode.InnerText.ToString());

                    }
                }

            }
        }

        public override double TimeSaloon()
        {
            if (RBtime == 0 || RC40time == 0)
            {
                return ((PartTime(FCtime, FCcount)) * 2 + 2 * (PartTime(FBtime, FBcount))) / 0.65;
            }
            else if (FCtime == 0 || FBcount == 0)
            {
                return (2 * (PartTime(RBtime, RBcount)) + PartTime(RC40time, RC40count)) / 0.35;
            }
            else
            {
                Double percent = (double)(RC40time / (RC40time + RC100time));
                return (PartTime(FCtime, FCcount)) * 2 + 2 * (PartTime(FBtime, FBcount)) + 2 * (PartTime(RBtime, RBcount)) + (1 - percent) * (RC100time / RC100count) + percent * (2 * RC40time / RC40count);
            }
        }

        public override int UniqueLines()
        {
            return LD.Count;
        }
    }
}
