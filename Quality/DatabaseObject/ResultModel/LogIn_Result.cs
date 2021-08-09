using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class LogIn_Result : ResultModelBase<LogIn_Result>
    {
        public Quality_Pass1 pass1 { get; set; }

        public List<Quality_Menu> Menus { get; set; }

        public List<string> Factorys { get; set; }

        public List<string> Lines { get; set; }
    }
}
