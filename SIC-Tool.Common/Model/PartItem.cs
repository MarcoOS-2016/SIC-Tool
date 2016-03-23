using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SIC_Tool.Common.Model
{
    public class PartItem
    {
        private string item;
        private string description;
        private string cc;
        private double partcost;

        public string Item
        {
            get { return item; }
            set { item = value; }
        }

        public string Description
        {
            get { return description; }
            set { description = value; }
        }

        public string CC
        {
            get { return cc; }
            set { cc = value; }
        }

        public double PartCost
        {
            get { return partcost; }
            set { partcost = value; }
        }
    }
}
