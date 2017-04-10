using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace LaborNeedsSchedulingFramework.Models
{
    public class Value
    {
        public object Date { get; set; }
        public object Weekday { get; set; }
        public object HourofDay { get; set; }
        public object TrafficOut { get; set; }


        public string payroll { get; set; }

    }
}