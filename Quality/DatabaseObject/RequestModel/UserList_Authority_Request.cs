using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.RequestModel
{
    public class UserList_Authority_Request
    {
        public string UserID { get; set; }
        public string Name { get; set; }
        public string Password { get; set; }
        public string Email { get; set; }

        public string Position { get; set; }

        public List<Module_Detail> DataList { get; set; }
    }
}
