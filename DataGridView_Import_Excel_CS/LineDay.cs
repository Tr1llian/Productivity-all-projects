using System;

namespace Productivity
{
    public class LineDay : IEquatable<LineDay>, IComparable<LineDay>
    {
        public int Name { get; set; }
        public string Date { get; set; }

        public LineDay(int line, string date)
        {
            Name = line;
            Date = date;
        }

        public bool Equals(LineDay other)
        {
            return Date == other.Date && Name == other.Name;
        }

        public int CompareTo(LineDay other)
        {
             if(Name>other.Name)
            {
                return 1;
            }
            else if( Name == other.Name)
            {
                if (DateTime.Parse(Date) <= DateTime.Parse(other.Date))
                {
                    return -1;
                }
                else if (DateTime.Parse(Date) >= DateTime.Parse(other.Date))
                {
                    return 1;
                }
                else 
                {
                    return 0;
                }

            }
            else 
            {
                return -1;
            }
        }
    }
}
