using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class CFTComments_ViewModel : ResultModelBase<CFTComments_Result>
    {
        public string QueryType { get; set; }
        public string OrderID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string SampleStage { get; set; }
        public string ReleasedBy { get; set; }

        /// <summary>
        /// for報表用，不同Sample Stage需要區分成不同報表
        /// </summary>
        public List<string> SampleStageList { get; set; }

        public string TempFileName{ get; set; }

    }

    public class CFTComments_Result
    {
        public string SampleStage { get; set; }
        public string CommentsCategory { get; set; }
        public string Comnments { get; set; }

        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string Article { get; set; }
        public string SizeCode { get; set; }
        public string Name { get; set; }
    }
}
