using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class GarmentTest_ViewModel
    {
        public string StyleID { get; set; }

        public string Brand { get; set; }

        public string Article { get; set; }

        public string Season { get; set; }

        public List<string> StyleID_Lsit { get; set; }

        public List<string> Brand_List { get; set; }

        public List<string> Season_List { get; set; }

        public List<string> Article_List { get; set; }
    }
}
