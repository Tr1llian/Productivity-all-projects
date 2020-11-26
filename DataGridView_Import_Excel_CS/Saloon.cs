using DocumentFormat.OpenXml.Office2010.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace Productivity
{
    abstract public class Saloon
    {
        public double FCtime, FBtime, RBtime, RCtime, RB40time, RB60time, RC40time, RC60time, SPtime, RB20time, RC100time, RB100time;
        public int FCcount, FBcount, RBcount, RCcount, RB40count, RB60count, RC40count, RC60count, SPcount, RB20count, RC100count, RB100count;

        public int lines = 0;
        public int days = 5;
        public List<LineDay> LD = new List<LineDay>();

        public abstract void InitLines();
        public abstract int UniqueLines();
        public string ProjectName { get; set; }

        public void UpdateLines(int b)
        {
            lines = b;
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
                // обходимо всі дочірні елементи 
                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    if(childnode.Attributes[0].Value==ProjectName.ToString())
                    {
                        childnode.InnerText = b.ToString();
                    }
                }
                xDoc.Save("Settings.xml");
            }
        }

        public void UpdateDays(int day)
        {
            days = day;
        }

        public double Coef { get; set; }
        public abstract void ParseExcel(DataRow row);
        public abstract double TimeSaloon();
        public abstract void CreateRow(ref DataRow row1);
        public abstract double GeneralCount();
        public abstract double GeneralTime();
        public abstract double AvgTime();

        public double PartTime(double a, double b)
        {
            if (a != 0 && b != 0)
            {
                return a / b;
            }
            else return 0;
        }
    }
}
