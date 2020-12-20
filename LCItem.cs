using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LCRanking
{
    public class LCItem
    {
        public string Name { get; set; }
        public int Point { get; set; }
        public int ItemId { get; set; }
        public double Deducted { get; set; }
        public int PassPoints { get; set; }
        public int Attendance { get; set; }
        public double Priority { get; set; }

    }
}
