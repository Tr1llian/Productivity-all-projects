using System;
using System.Data;

namespace Productivity
{
    public class SaloonG11 : Saloon
    {
        public SaloonG11(string name)
        {
            ProjectName = name;
        }

        public override double AvgTime()
        {
            double AllPcs = FCcount + FBcount + RC100count * 2 + RC40count + RBcount;
            double AllTime = FCtime + FBtime + RBtime + RCtime;
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
            RCtime = RC40time + RC100time;
            RCcount = RC100count + RC40count;
            row1["Проект"] = ProjectName;
            row1["Кількість чохлів"] = "\n FB = " + FBcount
                + "\n" + " FC = " + FCcount
                + "\n" + " RB = " + RBcount
                + "\n" + " RC40 = " + RC40count
                + "\n" + " RC100 = " + RC100count +
                "\n" + "Загальна кількість=" + GeneralCount() + "\n";
            row1["Загальний час"] = "\n FB time = " + FBtime
                + "\n" + " FC time = " + FCtime
                + "\n" + " RB time = " + RBtime
                + "\n" + " RC time = " + RCtime +
                "\n" + "Загальна час=" + GeneralTime() + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs= " + Math.Round(PartTime(FBtime, FBcount), 3)
                + "\n" + " FC time for pcs= " + Math.Round(PartTime(FCtime, FCcount), 3)
                + "\n" + " RB time for pcs= " + Math.Round(PartTime(RBtime, RBcount), 3)
                + "\n" + " RC time for pcs= " + Math.Round(PartTime(RCtime, RCcount), 3) + "\n";
            row1["Час на салон"] = Math.Round(TimeSaloon(), 3);
            row1["Кількість салонів"] = Math.Floor(GeneralCount() / Coef);
            row1["Середній час на одну штуку"] = Math.Round(AvgTime(), 3);
            row1["Коефіцієнт/кількість компонентів"] = Coef;
            row1["Кількість компонент помножено на середній на одну штуку"] = Math.Round(Coef * AvgTime(), 3);
            row1["Prod. sets planned"] = Math.Round(480 / (Coef * AvgTime()), 3);

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
            if (row[6].ToString().ToUpper().Contains("FC"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    FCcount += 1;
                }
                FCtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("FB"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    FBcount += 1;
                }
                FBtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("RC"))
            {
                if (row[6].ToString().ToUpper().Contains("RC100"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        RC100count += 1;
                    }
                    RC100time += Convert.ToInt16(row[7].ToString());
                }
                else
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        RC40count += 1;
                    }
                    RC40time += Convert.ToInt16(row[7].ToString());
                }


            }
            else if (row[6].ToString().ToUpper().Contains("RB"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    RBcount += 1;
                }
                RBtime += Convert.ToInt16(row[7].ToString());
            }
        }

        public override double TimeSaloon()
        {
            if (RBtime == 0 || RCtime == 0)
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
    }
}
