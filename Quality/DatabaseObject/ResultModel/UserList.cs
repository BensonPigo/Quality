using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class UserList : ResultModelBase<UserList_Browse>
    {
        public List<UserList_Browse> data { get; set; }
    }
}
