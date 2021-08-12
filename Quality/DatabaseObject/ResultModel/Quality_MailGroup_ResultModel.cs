using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class Quality_MailGroup_ResultModel : Quality_MailGroup
    {
        public bool Result { get; set; }

        public string ErrMsg { get; set; }
    }
}
