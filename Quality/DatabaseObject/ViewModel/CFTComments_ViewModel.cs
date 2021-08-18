using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class CFTComments_ViewModel
    {
        public string SampleStage { get; set; }
        public string CommentsCategory { get; set; }
        public string Comnments { get; set; }
    }

    public class CFTComments_where
    {
        public string OrderID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
    }
}
