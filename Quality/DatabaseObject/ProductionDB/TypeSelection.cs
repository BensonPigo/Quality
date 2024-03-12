using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class TypeSelection
    {
        [Required]
        public int VersionID { get; set; }

        [Required]
        public int Seq { get; set; }

        [Required]
        [StringLength(50)]
        public string Code { get; set; }

    }
}
