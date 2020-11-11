using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace Productivity
{
    public class SaloonQ3 : Saloon
    {
        public SaloonQ3(string name)
        {
            ProjectName = name;
            InitLines();
        }

        public override double AvgTime()
        {
            double AllPcs = FCcount + FBcount + RC40count+RC60count + RB40count + RB60count + RB20count;
            double AllTime = FCtime + FBtime + RB20time + RB40time + RB60time + RC40time+RC60time;
            if (AllPcs == 0)
            {
                return 0;
            }
            else
            {
                return AllTime / AllPcs;
            }
        }

        public override void CreateRow(ref DataRow row1)
        {
            row1["Проект"] = ProjectName;
            row1["Кількість чохлів"] = "\n FB = " + FBcount
                + "\n" + " FC = " + FCcount
                + "\n" + " RB = " + (RBcount + RB60count + RB40count)
                + "\n" + " RB20 = " + RB20count
                + "\n" + " RC40 = " + RC40count 
                +"\n" + " RC60 = " + RC60count +
                "\n" + "Загальна кількість=" + GeneralCount() + "\n";
            row1["Загальний час"] = "\n FB time = " + FBtime
                + "\n" + " FC time = " + FCtime
                + "\n" + " RB time = " + (RBtime + RB40time + RB60time)
                + "\n" + " RB20 time = " + RB20time
                + "\n" + " RC40 time = " + RC40time 
                  +"\n" + " RC60 time = " + RC60time +
                   "\n" + "Загальна час=" + GeneralTime() + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs= " + Math.Round(PartTime(FBtime, FBcount), 3)
                + "\n" + " FC time for pcs= " + Math.Round(PartTime(FCtime, FCcount), 3)
                + "\n" + " RB time for pcs= " + Math.Round(PartTime(RB60time + RB40time, RB60count + RB40count), 3)
                + "\n" + " RB20 time for pcs= " + Math.Round(PartTime(RB20time, RB20count), 3)
                + "\n" + " RC time for pcs= " + Math.Round(PartTime(RC40time + RC60time,RC40count + RC60count), 3) + "\n";
            row1["Час на салон"] = Math.Round(TimeSaloon(), 3);
            row1["Кількість салонів"] = Math.Floor(GeneralCount() / Coef);
            row1["Середній час на одну штуку"] = Math.Round(AvgTime(), 3);
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
            return FCcount + FBcount + RC40count+RC60count + RB60count + RB40count + RB20count;
        }

        public override double GeneralTime()
        {
            return FCtime + FBtime + RB20time + RB40time + RB60time + RC40time+RC60time;
        }

        public override void ParseExcel(DataRow row)
        {
            LineDay l = new LineDay(Convert.ToInt32(row[2].ToString()), row[0].ToString());
            if (Convert.ToInt32(row[3].ToString()) >= 5000000)
            {
                if (!LD.Contains(l))
                {
                    LD.Add(l);
                }
            }

            if (row[6].ToString().Contains("FC"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    FCcount += 1;
                }
                FCtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().Contains("FB"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    FBcount += 1;
                }
                FBtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().Contains("RB"))
            {
                if (row[6].ToString().Contains("RB60"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        RB60count += 1;
                    }
                    RB60time += Convert.ToInt16(row[7].ToString());
                }
                else if (row[6].ToString().Contains("RB40") || row[6].ToString().Contains("RB20"))
                {
                    if (row[6].ToString().Contains("RB40"))
                    {
                        if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                        {
                            RB40count += 1;
                        }
                        RB40time += Convert.ToInt16(row[7].ToString());
                    }
                    else if (row[6].ToString().Contains("RB20"))
                    {
                        if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                        {
                            RB20count += 1;
                        }
                        RB20time += Convert.ToInt16(row[7].ToString());
                    }
                }
            }
            else if (row[6].ToString().Contains("RC"))
            {
                if (row[6].ToString().Contains("RC40") || row[6].ToString().Contains("RC20"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        RC40count += 1;
                    }
                    RC40time += Convert.ToInt16(row[7].ToString());

                }
                else if (row[6].ToString().Contains("RC60"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        RC60count += 1;
                    }
                    RC60time += Convert.ToInt16(row[7].ToString());
                }
                else
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        RCcount += 1;
                    }
                    RCtime += Convert.ToInt16(row[7].ToString());
                }

            }
        }

        double RB20Coef(double rb20count,double saloons)
        {
            if(rb20count !=0)
            {
                return rb20count / saloons;
            }
            else
            {
                return 0;
            }
        }

        public override double TimeSaloon()
        {
            RBtime = RB40time + RB60time;
            RBcount = RB60count + RB40count;
            RCtime = RC40time + RC60time;
            RCcount = RC40count + RC60count;
 
            if (RB20count != 0 && RB20time != 0)
            {

                if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
                {
                    return 2.0 * PartTime(FCtime, FCcount) + 2.0 * PartTime(FBtime, FBcount) + 2.0 * PartTime(RB40time + RB60time, RB40count + RB60count) + 2.0 * PartTime(RC40time + RC60time, RC40count + RC60count) + RB20Coef(RB20count,(GeneralCount()/Coef))* PartTime(RB20time, RB20count);
                }
                if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount)) / 0.55 + PartTime(RB20time, RB20count);
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return (2 * RBtime / RCcount + 2 * PartTime(RCtime, RCcount)) / 0.45 + PartTime(RB20time, RB20count);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime == 0)
                {
                    return PartTime(FCtime, FCcount) / 0.1 + PartTime(RB20time, RB20count);
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (PartTime(FBtime, FBcount)) / 0.17 + PartTime(RB20time, RB20count);
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
                {
                    return (PartTime(RBtime, RBcount)) / 0.1 + PartTime(RB20time, RB20count);
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
                {
                    return (PartTime(RCtime, RCcount)) / 0.1 + PartTime(RB20time, RB20count);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return ((2 * PartTime(RBtime, RBcount) + 2 * PartTime(RCtime, RCcount)) + 2 * PartTime(FCtime, FCcount)) / 0.66 + PartTime(RB20time, RB20count); ;
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
                {
                    return ((2 * PartTime(RBtime, RBcount) + 2 * PartTime(FBtime, FBcount)) + 2 * (PartTime(FBtime, FBcount))) / 0.8 + PartTime(RB20time, RB20count);
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
                {
                    return (2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount)) / 0.55 / 2 + (PartTime(RCtime, RCcount)) / 0.1 / 2 + PartTime(RB20time, RB20count);
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
                {
                    return ((2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount)) / 0.55 / 2 + (PartTime(RBtime, RBcount)) / 0.1 / 2) + PartTime(RB20time, RB20count);
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
                {
                    return (double)(((PartTime(FBtime, FBcount)) / 0.1 / 2 + (PartTime(RCtime, RCcount)) / 0.1) / 2) + PartTime(RB20time, RB20count);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
                {
                    return (double)((PartTime(FCtime, FCcount) / 0.17 / 2 + (PartTime(RCtime, RCcount)) / 0.1) / 2) + PartTime(RB20time, RB20count);
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
                {
                    return (double)(((PartTime(FBtime, FBcount)) / 0.1 / 2 + (PartTime(RBtime, RBcount)) / 0.1) / 2) + PartTime(RB20time, RB20count);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
                {
                    return (double)((PartTime(FCtime, FCcount) / 0.17 / 2 + (PartTime(RBtime, RBcount)) / 0.1) / 2) + PartTime(RB20time, RB20count);
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
                {
                    return 2 * (PartTime(FCtime, FCcount) + PartTime(FBtime, FBcount) + PartTime(RBtime, RBcount) + PartTime(RCtime, RCcount));
                }
                if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount)) / 0.55;
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return (2 * RBtime / RCcount + 2 * PartTime(RCtime, RCcount)) / 0.45;
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime == 0)
                {
                    return PartTime(FCtime, FCcount) / 0.1;
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (PartTime(FBtime, FBcount)) / 0.17;
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
                {
                    return (PartTime(RBtime, RBcount)) / 0.1;
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
                {
                    return (PartTime(RCtime, RCcount)) / 0.1;
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return ((2 * PartTime(RBtime, RBcount) + 2 * PartTime(RCtime, RCcount)) + 2 * PartTime(FCtime, FCcount)) / 0.66;
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
                {
                    return ((2 * PartTime(RBtime, RBcount) + 2 * PartTime(FBtime, FBcount)) + 2 * (PartTime(FBtime, FBcount))) / 0.8;
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
                {
                    return (2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount)) / 0.55 / 2 + (PartTime(RCtime, RCcount)) / 0.1 / 2;
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
                {
                    return ((2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount)) / 0.55 / 2 + (PartTime(RBtime, RBcount)) / 0.1 / 2);
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
                {
                    return (double)(((PartTime(FBtime, FBcount)) / 0.1 / 2 + (PartTime(RCtime, RCcount)) / 0.1) / 2);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
                {
                    return (double)((PartTime(FCtime, FCcount) / 0.17 / 2 + (PartTime(RCtime, RCcount)) / 0.1) / 2);
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
                {
                    return (double)(((PartTime(FBtime, FBcount)) / 0.1 / 2 + (PartTime(RBtime, RBcount)) / 0.1) / 2);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
                {
                    return (double)((PartTime(FCtime, FCcount) / 0.17 / 2 + (PartTime(RBtime, RBcount)) / 0.1) / 2);
                }
                else
                {
                    return 0;
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
                    if (childnode.Name == "Q3")
                    {
                        lines = Convert.ToInt32(childnode.InnerText.ToString());

                    }
                }

            }
        }

        public override int UniqueLines()
        {
            return LD.Count;
        }
    }
}
