using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using DatabaseObject.ResultModel;

namespace DatabaseObject.RequestModel
{
    public class Quality_Position_Request
    {
        [Required]
        public string Position { get; set; }
        public string Factory { get; set; }

        [StringLength(100)]
        public string Description { get; set; }

        public bool IsAdmin { get; set; }
        public bool Junk { get; set; }

        public List<Module_Detail> DataList { get; set; }
    }
}
