using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class StyleList : ResultModelBase<StyleList>
    {
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string Description { get; set; }
        public string Phase { get; set; }
        public string SMR { get; set; }
        public string Handle { get; set; }
        public string SpecialMark { get; set; }
        public string SampleRFT { get; set; }

        public string MsgScript { get; set; }
    }
}
