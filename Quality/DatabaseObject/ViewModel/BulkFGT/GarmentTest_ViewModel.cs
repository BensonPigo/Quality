using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class GarmentTest_ViewModel : GarmentTest
    {
        public string WashName { get; set; }

        public string Teamwear { get; set; }

        public DateTime? MinSciDelivery { get; set; }

        public DateTime? MinBuyerDelivery { get; set; }

        public string GarmentTestAddName { get; set; }

        public string GarmentTestEditName { get; set; }

        public List<string> StyleID_Lsit { get; set; }

        public List<string> Brand_List { get; set; }

        public List<string> Season_List { get; set; }

        public List<string> Article_List { get; set; }

        public bool SaveResult { get; set; }

        public string ErrMsg { get; set; }

        public string Sender { get; set; }

        public string SendDate { get; set; }
    }
}
