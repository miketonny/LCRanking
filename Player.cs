using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LCRanking
{
    public class Player
    {
        public string Name { get; set; }
        public List<Item> Items { get; set; }
        //public List<LCRawItem> RawItems { get; set; }
    }

    public class Item
    {
        public int ItemId { get; set; }
        public int Priority { get; set; }
        public double Deducted { get; set; }
        public double Passes { get; set; }

        public double Attnd { get; set; }

        public double CalculatedPriority
        {
            get
            {
                return Math.Round(Priority + Passes + (Attnd * 0.1) + Deducted, 1);
            }             
        }

    }

    //public class LCRawItem
    //{
    //    public int ItemId { get; set; }
    //    public int Priority { get; set; }

    //}
}
