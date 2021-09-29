using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel.StyleManagement
{
    public class ExceptionFD_ViewModel
    {
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public string Article { get; set; }
        public string ExpectionFormStatus { get; set; }
        public DateTime? ExpectionFormDate { get; set; }
        public string ExpectionFormRemark { get; set; }
    }
}
