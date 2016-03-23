using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SIC_Tool.Common.Model
{
    public class SupplierItem
    {
        private string suppliername;
        private Int32 onwayqty;
        private Int32 gloviaqty;
        private Int32 betweenqty;
        private double onwayvalue;
        private double gloviavalue;
        private double betweenvalue;
        private double percentage;
        private string owner;

        public string SupplierName
        {
            get { return suppliername; }
            set { suppliername = value; }
        }

        public Int32 OnWayQty
        {
            get { return onwayqty; }
            set { onwayqty = value; }
        }

        public Int32 GloviaQty
        {
            get { return gloviaqty; }
            set { gloviaqty = value; }
        }

        public Int32 BetweenQty
        {
            get { return betweenqty; }
            set { betweenqty = value; }
        }

        public double OnWayValue
        {
            get { return onwayvalue; }
            set { onwayvalue = value; }
        }

        public double GloviaValue
        {
            get { return gloviavalue; }
            set { gloviavalue = value; }
        }

        public double BetweenValue
        {
            get { return betweenvalue; }
            set { betweenvalue = value; }
        }

        public double Percentage
        {
            get { return percentage; }
            set { percentage = value; }
        }

        public string Owner
        {
            get { return owner; }
            set { owner = value; }
        }
    }
}
