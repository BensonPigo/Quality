using DatabaseObject.ManufacturingExecutionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class ReworkList_ViewModel
    {
        public List<string> SPList { get; set; }

        public List<string> StyleList { get; set; }

        public List<string> ArticleList { get; set; }

        public List<string> SizeList { get; set; }

        public string SP { get; set; }
        public string Style { get; set; }
        public string Article { get; set; }
        public string Size { get; set; }

        public string FactoryID { get; set; }
        public string Line { get; set; }
        public string Status { get; set; }

        public RFT_Inspection rft_Inspection { get; set; }

        public List<ReworkList> ReworkList { get; set; }
    }
}
