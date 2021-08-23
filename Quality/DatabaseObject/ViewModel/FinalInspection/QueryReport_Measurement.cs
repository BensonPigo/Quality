using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace DatabaseObject.ViewModel.FinalInspection
{
    public class QueryReport_Measurement
    {
        /// <summary>
        /// Time 是 string EX:10:00,11:00
        /// </summary>
        public String Time { get; set; }
        public String Article { get; set; }
        public String SizeCode { get; set; }
        public String Location { get; set; }
        public String Description { get; set; }
        public String Tol2 { get; set; }
        public String Tol1 { get; set; }
        public String SizeSpec { get; set; }
        public String SizeSpec2 { get; set; }
        public String diff { get; set; }

    }
}
