using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class Module_Detail
    {
        public long MenuID { get; set; }
        public string ModuleName { get; set; }
        public string FunctionName { get; set; }
        public bool Used { get; set; }
    }
}
