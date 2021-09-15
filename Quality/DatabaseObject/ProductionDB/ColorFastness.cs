using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class ColorFastness
    {
        public string ID { get; set; }

        public string POID { get; set; }

        public decimal TestNo { get; set; }

        public DateTime? InspDate { get; set; }

        public string Article { get; set; }

        public string Result { get; set; }

        public string Status { get; set; }

        public string Inspector { get; set; }

        public string Remark { get; set; }

        public string addName { get; set; }

        public DateTime? addDate { get; set; }

        public string EditName { get; set; }

        public DateTime? EditDate { get; set; }

        public int? Temperature { get; set; }

        public int? Cycle { get; set; }

        public string Detergent { get; set; }

        public string Machine { get; set; }

        public string Drying { get; set; }

        public Byte[] TestBeforePicture { get; set; }

        public Byte[] TestAfterPicture { get; set; }

    }
}
