using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using DatabaseObject.ViewModel.BulkFGT;
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

        IList<GarmentTest_Detail_Shrinkage> Get_Shrinkage(Int64 ID, string No);

        IList<Garment_Detail_Spirality> Get_Spirality(Int64 ID, string No);

        IList<GarmentTest_Detail_Apperance_ViewModel> Get_Apperance(Int64 ID, string No);

        IList<GarmentTest_Detail_FGWT_ViewModel> Get_FGWT(Int64 ID, string No);

        IList<GarmentTest_Detail_FGPT_ViewModel> Get_FGPT(Int64 ID, string No);

        GarmentTest_ViewModel Save_GarmentTest(GarmentTest_ViewModel garmentTest_ViewModel, List<GarmentTest_Detail> detail);
    }
}
