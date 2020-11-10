using System;
using System.Data;

namespace Productivity
{
    public class SaloonSK38 : Saloon
    {
        public SaloonSK38(string name)
        {
            ProjectName = name;
        }

        public override double AvgTime()
        {
            double AllPcs = FCcount + FBcount + RCcount + RB40count + RB60count;
            double AllTime = FCtime + FBtime + RCtime + RB40time + RB60time;
            if (AllPcs == 0)
            {
                return 0.0;
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
                + "\n" + " RB40 = " + (RB40count)
                + "\n" + " RB60 = " + (RB60count)
                + "\n" + " RC = " + RCcount +
                "\n" + "Загальна кількість=" + GeneralCount() + "\n";
            row1["Загальний час"] = "\n FB time = " + FBtime
                + "\n" + " FC time = " + FCtime
                + "\n" + " RB40 time = " + (RB40time)
                + "\n" + " RB60 time = " + (RB60time)
                + "\n" + " RC time = " + RCtime +
                "\n" + "Загальна час=" + GeneralTime() + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs= " + Math.Round(PartTime(FBtime, FBcount), 3)
                + "\n" + " FC time for pcs= " + Math.Round(PartTime(FCtime, FCcount), 3)
                + "\n" + " RB40 time for pcs= " + Math.Round(PartTime(RB40time, RB40count), 3)
                + "\n" + " RB60 time for pcs= " + Math.Round(PartTime(RB60time, RB60count), 3)
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
            return FCcount + FBcount + RCcount + RB40count + RB60count;
        }

        public override double GeneralTime()
        {
            return FCtime + FBtime + RCtime + RB40time + RB60time;
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

            else if (row[6].ToString().ToUpper().Contains("RB"))
            {
                if (row[6].ToString().ToUpper().Contains("RB40"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        RB40count += 1;
                    }
                    RB40time += Convert.ToInt16(row[7].ToString());
                }
                if (row[6].ToString().ToUpper().Contains("RB60"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        RB60count += 1;
                    }
                    RB60time += Convert.ToInt16(row[7].ToString());
                }
            }

            else if (row[6].ToString().ToUpper().Contains("RC"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    RCcount += 1;
                }
                RCtime += Convert.ToInt16(row[7].ToString());
            }
        }

        public override double TimeSaloon()
        {
            RBtime += RB40time + RB60time;
            RBcount += RB60count + RB40count;

            if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
            {
                return 2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount) + 2 * PartTime(RBtime, RBcount) + PartTime(RCtime, RCcount);
            }
            else if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
            {
                return (2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount)) / 0.582;
            }
            else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
            {
                return (PartTime(RCtime, RCcount) + PartTime(RB40time, RB40count) + PartTime(RB60time, RB60count)) / 0.418;
            }
            else if (FCtime == 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
            {
                return (RCtime / RCcount) / 0.1852;
            }
            else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
            {
                return (PartTime(RB40time, RB40count) + PartTime(RB60time, RB60count)) / 0.2328;
            }
            else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime == 0)
            {
                return (PartTime(FCtime, FCcount)) / 0.0931;
            }
            else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
            {
                return (PartTime(FBtime, FBcount)) / 0.198;
            }
            else
            {
                return 0;
            }

            /*
            if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
            {
                return 2*(PartTime(FCtime / FCcount) + PartTime(FBtime ,FBcount) + PartTime(RBtime,RBcount) + PartTime(RCtime , RCcount)) ;
            }
            if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
            {
                return (2 * PartTime(FCtime / FCcount) + 2 * PartTime(FBtime ,FBcount)) / 0.65;
            }
            else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
            {
                return ( RBtime / RCcount +  PartTime(RCtime , RCcount)) / 0.35;
            }
            else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime == 0)
            {
                return (PartTime(FCtime / FCcount)) / 0.1;
            }
            else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
            {
                return (PartTime(FBtime ,FBcount)) / 0.17;
            }
            else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
            {
                return (RBtime / RBcount) / 0.1;
            }
            else if (FCtime == 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
            {
                return (PartTime(RCtime , RCcount)) / 0.1;
            }
            else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
            {
                return (RBtime / FCcount + PartTime(RCtime , RCcount)) / 0.35/2 + (PartTime(FCtime / FCcount)) / 0.1 / 2;
            }
            else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
            {
                return (2 * RBtime / FCcount + 2 * PartTime(FBtime ,FBcount)) / 0.35/2 + (PartTime(FBtime ,FBcount)) / 0.17 / 2;
            }
            else if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
            {
                return (2 * PartTime(FCtime / FCcount) + 2 * PartTime(FBtime ,FBcount)) / 0.65/2 + (PartTime(RCtime , RCcount)) / 0.1 / 2;
            }
            else if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
            {
                return ((2 * PartTime(FCtime / FCcount) + 2 * PartTime(FBtime ,FBcount)) / 0.65/2 + (RBtime / RBcount) / 0.1/2) ;
            }
            else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
            {
                return (double)(((PartTime(FBtime ,FBcount)) / 0.1/2 + (PartTime(RCtime , RCcount)) / 0.1)/2);
            }
            else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
            {
                return (double)(((PartTime(FCtime / FCcount)) / 0.17 / 2 + (PartTime(RCtime , RCcount)) / 0.1) / 2);
            }
            else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
            {
                return (double)(((PartTime(FBtime ,FBcount)) / 0.1 / 2 + (RBtime / RBcount) / 0.1) / 2);
            }
            else if (FCtime != 0 && FBtime == 0 && RBtime!=0 && RCtime == 0)
            {
                return (double)(((PartTime(FCtime / FCcount)) / 0.17 / 2 + (RBtime / RBcount) / 0.1) / 2);
            }
             
            else
            {
                return 0;
            }
            */
        }
    }
}
