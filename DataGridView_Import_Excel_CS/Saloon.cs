using System.Data;

namespace Productivity
{
    abstract public class Saloon
    {
        public double FCtime, FBtime, RBtime, RCtime, RB40time, RB60time, RC40time, RC60time, SPtime, RB20time, RC100time, RB100time;
        public int FCcount, FBcount, RBcount, RCcount, RB40count, RB60count, RC40count, RC60count, SPcount, RB20count, RC100count, RB100count;

        public string ProjectName { get; set; }
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
