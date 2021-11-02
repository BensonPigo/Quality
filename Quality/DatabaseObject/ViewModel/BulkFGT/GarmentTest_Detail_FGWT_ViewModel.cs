using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel
{
    public class GarmentTest_Detail_FGWT_ViewModel: GarmentTest_Detail_FGWT
    {
        public string Result { get; set; }

        /// <summary>
        /// 1 : BeforeWash、AfterWash 可編輯
        /// 2 : Scale  可編輯
        /// 3 : 都不可編輯
        /// </summary>
        public string EditType { get; set; }


        /// <summary>
        /// B : Bottom
        /// T : Top
        /// S : Top+Bottom
        /// </summary>
        public string LocationText { get; set; }

        public bool IsInPercentage { get; set; }
    }
}
