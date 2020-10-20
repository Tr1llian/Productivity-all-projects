using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace DataGridView_Import_Excel
{
    public class FormatRow
    {
        //Format row for project Q3
        public static void Q3row(Saloon car, ref DataRow row1)
        {
            row1["Проект"] = car.ProjectName;
            row1["Кількість чохлів"] = "\n FB = " + car.FBcount
                + "\n" + " FC= " + car.FCcount
                + "\n" + " RB= " + (car.RBcount + car.RB60count + car.RB40count)
                + "\n" + " RB20= " + car.RB20count
                + "\n" + " RC= " + car.RCcount + "\n";
            row1["Загальний час"] = "\n FB time = " + car.FBtime
                + "\n" + " FC time= " + car.FCtime
                + "\n" + " RB time= " + (car.RBtime + car.RB40time + car.RB60time)
                + "\n" + " RB20 time= " + car.RB20time
                + "\n" + " RC time= " + car.RCtime + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs= " + Math.Round(car.PartTime(car.FBtime, car.FBcount), 3)
                + "\n" + " FC time for pcs= " + Math.Round(car.PartTime(car.FCtime, car.FCcount), 3)
                + "\n" + " RB time for pcs= " + Math.Round(car.PartTime(car.RB60time + car.RB40time, car.RB60count + car.RB40count), 3)
                + "\n" + " RB20 time for pcs= " + Math.Round(car.PartTime(car.RB20time, car.RB20count), 3)
                + "\n" + " RC time for pcs= " + Math.Round(car.PartTime(car.RCtime + car.RC40time + car.RC60time, car.RCcount + car.RC40count + car.RC60count), 3) + "\n";
            row1["Час на салон"] = Math.Round(car.TimeSaloon(), 3);
            row1["Кількість салонів"] = Math.Floor(car.GeneralCount() / car.Coef);
            row1["Середній час на одну штуку"] = Math.Round(car.AvgTime(), 3);
        }

        //Format row for G11
        public static void G11row(Saloon car , ref DataRow row1)
        {
            double RCmiddle = 0;
            RCmiddle = (car.PartTime(car.RC40time, car.RC40count) + car.PartTime(car.RC100time, car.RC100count)) / 2;
            car.RCtime = car.RC40time + car.RC100time;
            car.RCcount = car.RC100count + car.RC40count;
            row1["Проект"] = car.ProjectName;
            row1["Кількість чохлів"] = "\n FB = " + car.FBcount
                + "\n" + " FC= " + car.FCcount
                + "\n" + " RB= " + car.RBcount
                + "\n" + " RC= " + car.RCcount + "\n";
            row1["Загальний час"] = "\n FB time = " + car.FBtime
                + "\n" + " FC time= " + car.FCtime
                + "\n" + " RB time= " + car.RBtime
                + "\n" + " RC time= " + car.RCtime + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs= " + Math.Round(car.PartTime(car.FBtime, car.FBcount), 3)
                + "\n" + " FC time for pcs= " + Math.Round(car.PartTime(car.FCtime, car.FCcount), 3)
                + "\n" + " RB time for pcs= " + Math.Round(car.PartTime(car.RBtime, car.RBcount), 3)
                + "\n" + " RC time for pcs= " + Math.Round(car.PartTime(car.RCtime, car.RCcount), 3) + "\n";
            row1["Час на салон"] = Math.Round(car.TimeSaloonBMW(), 3);
            row1["Кількість салонів"] = Math.Floor(car.GeneralCount() / car.Coef);
            row1["Середній час на одну штуку"] = Math.Round(car.AvgTime(), 3);
        }

        //Format row for G3
        public static void G3row(Saloon car, ref DataRow row1)
        {
            double RCmiddle = 0;
            RCmiddle = (car.PartTime(car.RC40time, car.RC40count) + car.PartTime(car.RC100time, car.RC100count)) / 2;
            car.RCtime = car.RC40time + car.RC100time;
            car.RCcount = car.RC100count + car.RC40count;
            row1["Проект"] = car.ProjectName;
            row1["Кількість чохлів"] = "\n FB = " + car.FBcount
                + "\n" + " FC= " + car.FCcount
                + "\n" + " RB= " + car.RBcount
                + "\n" + " RC= " + car.RCcount + "\n";
            row1["Загальний час"] = "\n FB time = " + car.FBtime
                + "\n" + " FC time= " + car.FCtime
                + "\n" + " RB time= " + car.RBtime
                + "\n" + " RC time= " + car.RCtime + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs= " + Math.Round(car.PartTime(car.FBtime, car.FBcount), 3)
                + "\n" + " FC time for pcs= " + Math.Round(car.PartTime(car.FCtime, car.FCcount), 3)
                + "\n" + " RB time for pcs= " + Math.Round(car.PartTime(car.RBtime, car.RBcount), 3)
                + "\n" + " RC time for pcs= " + Math.Round(car.PartTime(car.RCtime, car.RCcount), 3) + "\n";
            row1["Час на салон"] = Math.Round(car.TimeSaloonBMW(), 3);
            row1["Кількість салонів"] = Math.Floor(car.GeneralCount() / car.Coef);
            row1["Середній час на одну штуку"] = Math.Round(car.AvgTime(), 3);
        }

        //Format row for BR223 
        public static void BR223row(Saloon car, ref DataRow row1)
        {
            row1["Проект"] = car.ProjectName;
            row1["Кількість чохлів"] = "\n FB = " + car.FBcount
                + "\n" + " FC = " + car.FCcount
                 + "\n" + " VST = " + car.VSTcount
                + "\n" + " RB = " + (car.RBcount + car.RB60count + car.RB40count)
                + "\n" + " RC = " + car.RCcount + "\n";
            row1["Загальний час"] = "\n FB time = " + car.FBtime
                + "\n" + " FC time = " + car.FCtime
                + "\n" + " VST time = " + car.VSTtime
                + "\n" + " RB time = " + (car.RBtime + car.RB40time + car.RB60time)
                + "\n" + " RC time = " + car.RCtime + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs = " + Math.Round(car.PartTime(car.FBtime, car.FBcount), 3)
                + "\n" + " FC time for pcs = " + Math.Round(car.PartTime(car.FCtime, car.FCcount), 3)
                 + "\n" + " VST time for pcs = " + Math.Round(car.PartTime(car.VSTtime, car.VSTcount), 3)
                + "\n" + " RB time for pcs = " + Math.Round(car.PartTime(car.RBtime, car.RBcount), 3)
                + "\n" + " RC time for pcs = " + Math.Round(car.PartTime(car.RCtime, car.RCcount), 3) + "\n";
            row1["Час на салон"] = Math.Round(car.TimeSaloonBR223(), 3);
            row1["Кількість салонів"] = Math.Floor(car.GeneralCountBR223() / car.Coef);
            row1["Середній час на одну штуку"] = Math.Round(car.AvgTimeBR223(), 3);
        }

        //Format row for SK38
        public static void SK38row(Saloon car , ref DataRow row1)
        {
            row1["Проект"] = car.ProjectName;
            row1["Кількість чохлів"] = "\n FB = " + car.FBcount
                + "\n" + " FC= " + car.FCcount
                + "\n" + " RB40= " + (car.RB40count)
                + "\n" + " RB60= " + (car.RB60count)
                + "\n" + " RC= " + car.RCcount + "\n";
            row1["Загальний час"] = "\n FB time = " + car.FBtime
                + "\n" + " FC time= " + car.FCtime
                + "\n" + " RB40 time= " + (car.RB40time)
                + "\n" + " RB60 time= " + (car.RB60time)
                + "\n" + " RC time= " + car.RCtime + "\n";
            row1["Час на одну штуку"] = "\n FB time for pcs= " + Math.Round(car.PartTime(car.FBtime, car.FBcount), 3)
                + "\n" + " FC time for pcs= " + Math.Round(car.PartTime(car.FCtime, car.FCcount), 3)
                + "\n" + " RB40 time for pcs= " + Math.Round(car.PartTime(car.RB40time, car.RB40count), 3)
                + "\n" + " RB60 time for pcs= " + Math.Round(car.PartTime(car.RB60time, car.RB60count), 3)
                + "\n" + " RC time for pcs= " + Math.Round(car.PartTime(car.RCtime, car.RCcount), 3) + "\n";
            row1["Час на салон"] = Math.Round(car.TimeSaloonSK38(), 3);
            row1["Кількість салонів"] = Math.Floor(car.GeneralCount() / car.Coef);
            row1["Середній час на одну штуку"] = Math.Round(car.AvgTime(), 3);

        }
    }
}
