using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
  public class GarmentTest_Detail_FGPT_ViewModel: GarmentTest_Detail_FGPT
    {
        public string Result { get; set; }

        /// <summary>
        /// B : Bottom
        /// T : Top
        /// S : Top+Bottom
        /// </summary>
        public string LocationText { get; set; }
    }
}
