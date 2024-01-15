using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class UserList_Browse
    {
        public bool Select { get; set; }
        public string UserID { get; set; }
        public string Name { get; set; }
        public string Position { get; set; }
        public string Pivot88UserName { get; set; }
        public bool Junk { get; set; }
    }
}
