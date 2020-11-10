using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Productivity
{
    public class SaloonBR223 : Saloon
    {
        public SaloonBR223(string name)
        {
            ProjectName = name;
        }

        public override double AvgTime()
        {
            double AllCount = FCcount + FBcount + RC40count + 2 * RC100count + RB40count + RB100count * 2;
            double AllTime = FCtime + FBtime + RBtime + RCtime + SPtime;
            if (AllCount == 0)
            {
                return 0;
            }
            else
            {
                return AllTime / AllCount;
            }
        }

        public override void CreateRow(ref DataRow row1)
        {
            row1["Проект"] = ProjectName;
            row1["Кількість чохлів"] = "\n FB = " + FBcount
                + "\n" + " FC = " + FCcount
                + "\n" + " RB40 (FES) = " + RB40count
                  + "\n" + " RB100 (FSS)= " + RB100count
                + "\n" + " RC40 (FES)= " + RC40count
                + "\n" + " RC100 (FSS)= " + RC100count
                 + "\n" + " Small projects = " + SPcount
                + "\n" + "Загальна кількість=" + GeneralCount() + "\n";
            row1["Загальний час"] = "\n FB time = " + FBtime
                + "\n" + " FC time = " + FCtime

                + "\n" + " RB time = " + (RBtime + RB40time + RB60time)
                + "\n" + " RC time = " + RCtime +
                "\n" + " Small projects time = " + SPtime
                + "\n" + "Загальний час=" + GeneralTime() + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs = " + Math.Round(PartTime(FBtime, FBcount), 3)
                + "\n" + " FC time for pcs = " + Math.Round(PartTime(FCtime, FCcount), 3)
                + "\n" + " RB time for pcs = " + Math.Round(PartTime(RBtime, RB100count * 2 + RB40count), 3)
                + "\n" + " RC time for pcs = " + Math.Round(PartTime(RCtime, RC40count + RC100count * 2), 3)
                + "\n" + " Small projects time for pcs = " + Math.Round(PartTime(SPtime, SPcount), 3);
            row1["Час на салон"] = Math.Round(TimeSaloon(), 3);
            row1["Кількість салонів"] = Math.Floor(GeneralCount() / Coef);
            row1["Середній час на одну штуку"] = Math.Round(AvgTime(), 3);
            row1["Коефіцієнт/кількість компонентів"] = Coef;
            row1["Кількість компонент помножено на середній на одну штуку"] = Math.Round(Coef * AvgTime(), 3);
            row1["Prod. sets planned"] = Math.Round(480 / (Coef * AvgTime()), 3);
        }

        public override double GeneralCount()
        {
            return FCcount + FBcount + RC60count + RB40count + RB100count * 2 + RC40count + RC100count * 2;
        }

        public override double GeneralTime()
        {
            return FCtime + FBtime + RBtime + RCtime + SPtime;
        }

        public override void ParseExcel(DataRow row)
        {
            if (row[6].ToString().ToUpper().Contains("VST") || row[6].ToString().ToUpper().Contains("MITTE") || row[6].ToString().ToUpper().Contains("M-TEIL") || row[6].ToString().Contains("Motorschutzabdeckung") || row[6].ToString().Contains("Sichtschutz"))
            {
                SPtime += Convert.ToInt16(row[7].ToString());
                SPcount++;
            }
            //Console.WriteLine(row[6].ToString());
            if (row[6].ToString().ToUpper().Contains("FAKI"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    FCcount += 1;
                }
                FCtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("FALE"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    FBcount += 1;
                }
                FBtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("FOLE"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    if ((row[6].ToString().ToUpper().Contains("FSS")))
                    {
                        RB100count += 1;
                    }
                    else if ((row[6].ToString().ToUpper().Contains("FES")))
                    {
                        RB40count += 1;
                    }
                }
                if (row[6].ToString().ToUpper().Contains("FES") || (row[6].ToString().ToUpper().Contains("FSS")))
                {
                    RBtime += Convert.ToInt16(row[7].ToString());
                }
            }

            else if (row[6].ToString().ToUpper().Contains("FOKI"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    if (row[6].ToString().ToUpper().Contains("FOKI") && (row[6].ToString().ToUpper().Contains("FSS")))
                    {
                        RC100count += 1;
                    }
                    if (row[6].ToString().ToUpper().Contains("FOKI") && (row[6].ToString().ToUpper().Contains("FES")))
                    {
                        RC40count += 1;
                    }

                }
                if (row[6].ToString().ToUpper().Contains("FES") || (row[6].ToString().ToUpper().Contains("FSS")))
                {
                    RCtime += Convert.ToInt16(row[7].ToString());
                }
            }
        }

        public int SaloonCountBR223()
        {
            return Convert.ToInt32(Math.Floor(GeneralCount() / Coef));
        }

        public override double TimeSaloon()
        {
            RBtime += RB40time + RB60time;
            RBcount = RB60count + RB40count;
            RCcount = RC40count + RC100count;
            Console.WriteLine(PartTime(SPcount, SaloonCountBR223()) * PartTime(SPtime, SPcount));
            if (SPtime != 0 && SPcount != 0)
            {
                if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
                {
                    return 2.0 * PartTime(FCtime, FCcount) + 2.0 * PartTime(FBtime, FBcount) + 2.0 * PartTime(RBtime, RB40count + RB100count * 2) + 2.0 * PartTime(RCtime, RC40count + RC100count * 2) + PartTime(SPcount, SaloonCountBR223()) * PartTime(SPtime, SPcount);
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount)) / 0.4656 + PartTime(SPtime, SPcount);
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return (2 * PartTime(RBtime, RBcount) + 2 * PartTime(RCtime, RCcount)) / 0.5346 + PartTime(SPtime, SPcount);
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
                    return 2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount) + 2 * PartTime(RBtime, RBcount) + 2 * PartTime(RCtime, RCcount);
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (2 * PartTime(FCtime, FCcount) + 2 * PartTime(FBtime, FBcount)) / 0.4656;
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return (2 * PartTime(RBtime, RBcount) + 2 * PartTime(RCtime, RCcount)) / 0.5346;
                }
                else
                {
                    return 0;
                }
            }
            /*
            if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
            {
                return 2*(PartTime(FCtime , FCcount) + PartTime(FBtime , FBcount) + RBtime/RBcount + PartTime(RCtime , RCcount)) ;
            }
            if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
            {
                return (2 * PartTime(FCtime , FCcount) + 2 * PartTime(FBtime , FBcount)) / 0.65;
            }
            else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
            {
                return ( RBtime / RCcount +  PartTime(RCtime , RCcount)) / 0.35;
            }
            else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime == 0)
            {
                return (PartTime(FCtime , FCcount)) / 0.1;
            }
            else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
            {
                return (PartTime(FBtime , FBcount)) / 0.17;
            }
            else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
            {
                return (PartTime(RBtime , RBcount)) / 0.1;
            }
            else if (FCtime == 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
            {
                return (PartTime(RCtime , RCcount)) / 0.1;
            }
            else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
            {
                return (RBtime / FCcount + PartTime(RCtime , RCcount)) / 0.35/2 + (PartTime(FCtime , FCcount)) / 0.1 / 2;
            }
            else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
            {
                return (2 * RBtime / FCcount + 2 * PartTime(FBtime , FBcount)) / 0.35/2 + (PartTime(FBtime , FBcount)) / 0.17 / 2;
            }
            else if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
            {
                return (2 * PartTime(FCtime , FCcount) + 2 * PartTime(FBtime , FBcount)) / 0.65/2 + (PartTime(RCtime , RCcount)) / 0.1 / 2;
            }
            else if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
            {
                return ((2 * PartTime(FCtime , FCcount) + 2 * PartTime(FBtime , FBcount)) / 0.65/2 + (PartTime(RBtime , RBcount)) / 0.1/2) ;
            }
            else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
            {
                return (double)(((PartTime(FBtime , FBcount)) / 0.1/2 + (PartTime(RCtime , RCcount)) / 0.1)/2);
            }
            else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
            {
                return (double)(((PartTime(FCtime , FCcount)) / 0.17 / 2 + (PartTime(RCtime , RCcount)) / 0.1) / 2);
            }
            else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
            {
                return (double)(((PartTime(FBtime , FBcount)) / 0.1 / 2 + (PartTime(RBtime , RBcount)) / 0.1) / 2);
            }
            else if (FCtime != 0 && FBtime == 0 && RBtime!=0 && RCtime == 0)
            {
                return (double)(((PartTime(FCtime , FCcount)) / 0.17 / 2 + (PartTime(RBtime , RBcount)) / 0.1) / 2);
            }
             
            else
            {
                return 0;
            }
            */
        }
    }
}
