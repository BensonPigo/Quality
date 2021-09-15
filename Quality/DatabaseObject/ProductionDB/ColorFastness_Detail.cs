using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{

    public class ColorFastness_Detail
    {
        public string ID { get; set; }

        public string ColorFastnessGroup { get; set; }

        public string SEQ1 { get; set; }

        public string SEQ2 { get; set; }

        public string Roll { get; set; }

        public string Dyelot { get; set; }

        public string Result { get; set; }

        public string changeScale { get; set; }

        public string StainingScale { get; set; }

        public string Remark { get; set; }

        public string AddName { get; set; }

        public DateTime? AddDate { get; set; }

        public string EditName { get; set; }

        public DateTime? EditDate { get; set; }

        public DateTime? SubmitDate { get; set; }

        public string ResultChange { get; set; }

        public string ResultStain { get; set; }
    }
}
