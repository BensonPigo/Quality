using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel.EtoEFlowChart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ViewModel.StyleManagement
{
    public class EtoEFlowChart_ViewModel
    {
        public EtoEFlowChart_Request Request { get; set; }

        public bool ExecuteResult { get; set; }
        public string ErrorMessage { get; set; }

        public Development Development { get; set; }
        public Production Production { get; set; }
        public Warehouse Warehouse { get; set; }
        public DatabaseObject.ResultModel.EtoEFlowChart.FinalInspection FinalInspection { get; set; }

    }
}
