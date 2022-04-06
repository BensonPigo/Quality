using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class InspectionType
    {
        /// <summary>ID</summary>
        public Int64? ID { get; set; }

        /// <summary>Category</summary>
        public string Category { get; set; }

        /// <summary>Neme</summary>
        public string Neme { get; set; }

        /// <summary>Maximun</summary>
        public decimal Maximun { get; set; }

        /// <summary>Minimum</summary>
        public decimal Minimum { get; set; }

        /// <summary>MaxEqual</summary>
        public bool MaxEqual { get; set; }

        /// <summary>MinEqual</summary>
        public bool MinEqual { get; set; }

        /// <summary>UserCustomize</summary>
        public bool UserCustomize { get; set; }

        /// <summary>Comment</summary>
        public string Comment { get; set; }
    }
}
