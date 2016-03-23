using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SIC_Tool.Common.Model
{
    public class AgingItem
    {
        private string suppliername;
        private double kpigoal;
        private double days1;
        private double days2;
        private double days3;
        private double days4;
        private double days5;
        private double forwarder;
        private double onwayvalue;
        private double agingtotalcost;

        public string SupplierName
        {
            get { return suppliername; }
            set { suppliername = value; }
        }

        public double KPIGoal
        {
            get { return kpigoal; }
            set { kpigoal = value; }
        }

        public double Days1
        {
            get { return days1; }
            set { days1 = value; }
        }

        public double Days2
        {
            get { return days2; }
            set { days2 = value; }
        }

        public double Days3
        {
            get { return days3; }
            set { days3 = value; }
        }

        public double Days4
        {
            get { return days4; }
            set { days4 = value; }
        }

        public double Days5
        {
            get { return days5; }
            set { days5 = value; }
        }

        public double Forwarder
        {
            get { return forwarder; }
            set { forwarder = value; }
        }

        public double OnWayValue
        {
            get { return onwayvalue; }
            set { onwayvalue = value; }
        }

        public double AgingTotalCost
        {
            get { return agingtotalcost; }
            set { agingtotalcost = value; }
        }
    }
}
