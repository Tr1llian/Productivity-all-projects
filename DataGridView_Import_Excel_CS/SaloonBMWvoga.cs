using System;
using System.Data;

namespace Productivity
{
    class SaloonBMWvoga : Saloon
    {
        public SaloonBMWvoga(Saloon G11, Saloon G3, string name)
        {
            FBcount = G11.FBcount + G3.FBcount;
            FCcount = G11.FCcount + G3.FCcount;
            FBtime = G11.FBtime + G3.FBtime;
            FCtime = G11.FCtime + G11.FCtime;
            ProjectName = name;
        }

        public override double AvgTime()
        {
            return Math.Round(((double)(FCtime + FBtime) / (double)(FCcount + FBcount)), 3);
        }

        public override void CreateRow(ref DataRow row1)
        {
            row1["Проект"] = "BMW Voga";
            row1["Кількість чохлів"] = "\n FB = " + FBcount
                + "\n" + " FC = " + FCcount +
                "\n" + "Загальна кількість=" + GeneralCount() + "\n";
            row1["Загальний час"] = "\n FB time = " + FBtime
                + "\n" + " FC time = " + FCtime +
                "\n" + "Загальна час=" + GeneralTime() + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs= " + Math.Round(PartTime(FCtime, FBcount), 3)
                + "\n" + " FC time for pcs= " + Math.Round(PartTime(FCtime, FCcount), 3) + "\n";
            row1["Час на салон"] = Math.Round(TimeSaloon());
            row1["Кількість салонів"] = Math.Floor((FBcount + FCcount) / Coef);
            row1["Середній час на одну штуку"] = Math.Round(((double)(FCtime + FBtime) / (double)(FCcount + FBcount)), 3);
            row1["Коефіцієнт/кількість компонентів"] = Coef;
            row1["Кількість компонент помножено на середній на одну штуку"] = Math.Round(Coef * AvgTime(), 3);
            row1["Prod. sets planned"] = Math.Round(480 / (Coef * AvgTime()), 3);
        }
            
        public override double GeneralCount()
        {
            return FCcount + FBcount + RCcount + RBcount;
        }

        public override double GeneralTime()
        {
            return FBtime + FCtime + RBtime + RCtime;
        }

        public override void ParseExcel(DataRow row)
        {
            Console.WriteLine("Not needed to parse");
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
    }
}
