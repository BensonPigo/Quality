using DatabaseObject.ViewModel.SampleRFT;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Service.SampleRFT
{
    public class InspectionDefectListService
    {
        public InspectionDefectListProvider _Provider { get; set; }

        public InspectionDefectList_ViewModel GetData(InspectionDefectList_ViewModel Req)
        {
            InspectionDefectList_ViewModel result = new InspectionDefectList_ViewModel()
            {
                OrderID = Req.OrderID,
                DataList = new List<InspectionDefectList_Result>(),
            };
            try
            {
                _Provider = new InspectionDefectListProvider(Common.ProductionDataAccessLayer);

                var list = _Provider.GetData(Req.OrderID).ToList();

                if (list.Any())
                {
                    result.SampleStage = list.FirstOrDefault().SampleStage;
                    result.StyleID = list.FirstOrDefault().StyleID;
                    result.SeasonID = list.FirstOrDefault().SeasonID;
                    if (list.Where(o => !string.IsNullOrEmpty(o.Article)).Any())
                    {
                        result.DataList = list;
                    }
                }

                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = $@"msg.WithInfo('{ex.Message}');";
            }

            return result;
        }
    }
}
