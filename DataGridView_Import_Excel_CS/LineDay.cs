using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Productivity
{
    public class LineDay : IEquatable<LineDay>
    {
        int Name { get; set; }
        string Date { get; set; }

        public LineDay(int line, string date)
        {
            Name = line;
            Date = date;
        }

        public bool Equals(LineDay other)
        {
            return Date == other.Date && Name == other.Name;
        }
    }
}
