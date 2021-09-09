using ADOHelper.Template.MSSQL;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using ToolKit;
using ManufacturingExecutionDataAccessLayer.Interface;
using DatabaseObject.ManufacturingExecutionDB;
using System.Linq;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using DatabaseObject.ResultModel;
using BusinessLogicLayer.Interface.SampleRFT;

namespace BusinessLogicLayer.SampleRFT.Service
{
    public class BACriteriaService : IBACriteriaService
    {
        public IBACriteriaProvider _BACriteriaProvider { get; set; }

        public BACriteria_ViewModel Get_BACriteria_Result(BACriteria_ViewModel Req)
        {
            BACriteria_ViewModel result = new BACriteria_ViewModel();

            try
            {
                _BACriteriaProvider = new BACriteriaProvider(Common.ManufacturingExecutionDataAccessLayer);

                List<BACriteria_Result> list = _BACriteriaProvider.Get_BACriteria_Result(Req.OrderID, Req.StyleID, Req.BrandID, Req.SeasonID).ToList();

                int SumBAProduct = list.Sum(o => o.BAProduct);
                int SumInspectedQty = list.Sum(o => o.InspectedQty);

                if (SumInspectedQty == 0)
                {
                    result.SummaryBACriteria = 0;
                }
                else
                {
                    result.SummaryBACriteria = Convert.ToDecimal(1.0 * SumBAProduct / SumInspectedQty * 5);
                    result.SummaryBACriteria = Convert.ToDecimal(8.7);
                }
                result.DataList = list;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message;
            }


            return result;
        }
    }
}
