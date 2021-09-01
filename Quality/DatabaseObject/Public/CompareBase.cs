using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.Public
{
    public enum CompareStateType
    {
        Add,
        Edit,
        Delete,
        None
    }

    public class CompareBase
    {
        public CompareStateType StateType { get; set; } = CompareStateType.None;
    }
}
