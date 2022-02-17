using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.RequestModel
{
    public class GarmentTest_Request
    {
        public string Brand { get; set; } 

        public string Season { get;set; }

        public string Style { get; set; } 

        public string Article { get; set; } 

        public string Factory { get; set; }

        public string MDivisionid { get; set; }

        public string ID { get; set; }
    }

    public class P04Data
    {
        /// <inheritdoc/>
        public DateTime? DateSubmit { get; set; }

        /// <inheritdoc/>
        public decimal? NumArriveQty { get; set; }

        /// <inheritdoc/>
        public string TxtSize { get; set; }

        /// <inheritdoc/>
        public bool? RdbtnLine { get; set; }

        /// <inheritdoc/>
        public bool? RdbtnTumble { get; set; }

        /// <inheritdoc/>
        public bool? RdbtnHand { get; set; }

        /// <inheritdoc/>
        public string ComboTemperature { get; set; }

        /// <inheritdoc/>
        public string ComboMachineModel { get; set; }

        /// <inheritdoc/>
        public string TxtFibreComposition { get; set; }

        /// <inheritdoc/>
        public string ComboNeck { get; set; }

        /// <inheritdoc/>
        public string NumTwisTingBottom { get; set; }

        /// <inheritdoc/>
        public decimal? NumBottomS1 { get; set; }

        /// <inheritdoc/>
        public decimal? NumBottomL { get; set; }

        /// <inheritdoc/>
        public string NumTwisTingOuter { get; set; }

        /// <inheritdoc/>
        public decimal? NumOuterS1 { get; set; }

        /// <inheritdoc/>
        public decimal? NumOuterS2 { get; set; }

        /// <inheritdoc/>
        public decimal? NumOuterL { get; set; }

        /// <inheritdoc/>
        public string NumTwisTingInner { get; set; }

        /// <inheritdoc/>
        public decimal? NumInnerS1 { get; set; }

        /// <inheritdoc/>
        public decimal? NumInnerS2 { get; set; }

        /// <inheritdoc/>
        public decimal? NumInnerL { get; set; }

        /// <inheritdoc/>
        public string NumTwisTingTop { get; set; }

        /// <inheritdoc/>
        public decimal? NumTopS1 { get; set; }

        /// <inheritdoc/>
        public decimal? NumTopS2 { get; set; }

        /// <inheritdoc/>
        public decimal? NumTopL { get; set; }

        /// <inheritdoc/>
        public string TxtLotoFactory { get; set; }
        public string Remark { get; set; }
    }
}
