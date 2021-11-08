using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class UserList_Authority : ResultModelBase<Module_Detail>
    {
        public string UserID { get; set; }
        public string Name { get; set; }
        public string Password { get; set; }
        public string Email { get; set; }
        public string Position { get; set; }
        public string Pivot88UserName { get; set; }
    }
}
