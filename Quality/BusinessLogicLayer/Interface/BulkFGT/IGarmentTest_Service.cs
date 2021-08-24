using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BusinessLogicLayer.Service.BulkFGT.GarmentTest_Service;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IGarmentTest_Service
    {
        GarmentTest_ViewModel GetSelectItemData(GarmentTest_ViewModel garmentTest_ViewModel, SelectType type);

        GarmentTest_Result GetGarmentTest(GarmentTest_ViewModel garmentTest_ViewModel);

        List<string> Get_SizeCode(string OrderID, string Article);
    }
}
