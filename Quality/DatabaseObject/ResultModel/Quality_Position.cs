using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace DatabaseObject.ResultModel
{
    public class Quality_Position : ResultModelBase<Module_Detail>
    {
        [Required]
        public string Position { get; set; }
        public bool IsAdmin { get; set; }


        public string Description { get; set; }
        public bool Junk { get; set; }
    }
}
