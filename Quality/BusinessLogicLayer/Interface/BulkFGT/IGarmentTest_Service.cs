using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
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

        GarmentTest_Result GetGarmentTest(GarmentTest_Request garmentTest_ViewModel);

        List<string> Get_SizeCode(string OrderID, string Article);

        List<string> Get_SizeCode(string StyleID, string SeasonID, string BrandID);

        bool CheckOrderID(string OrderID, string BrandID, string SeasonID, string StyleID);

        GarmentTest_ViewModel SendMail(string ID, string No, string UserID);

        GarmentTest_ViewModel ReceiveMail(string ID, string No, string UserID);

        IList<GarmentTest_Detail_Shrinkage> Get_Shrinkage(string ID, string No);

        IList<Garment_Detail_Spirality> Get_Spirality(string ID, string No);

        IList<GarmentTest_Detail_Apperance_ViewModel> Get_Apperance(string ID, string No);

        IList<GarmentTest_Detail_FGWT_ViewModel> Get_FGWT(string ID, string No);

        IList<GarmentTest_Detail_FGPT_ViewModel> Get_FGPT(string ID, string No);

        GarmentTest_ViewModel Save_GarmentTest(GarmentTest_ViewModel garmentTest_ViewModel, List<GarmentTest_Detail> detail, string UserID);

        GarmentTest_Result Generate_FGWT(GarmentTest_ViewModel Main, GarmentTest_Detail_ViewModel Detail);

        GarmentTest_ViewModel DeleteDetail(string ID, string No);

        GarmentTest_Detail_Result Save_GarmentTestDetail(GarmentTest_Detail_Result source);

        string Get_LastResult(string ID);

        GarmentTest_Detail_Result Get_All_Detail(string ID, string No);

        GarmentTest_ViewModel Get_Main(string ID);

        GarmentTest_Detail_ViewModel Get_Detail(string ID, string No);

        List<string> Get_Scales();

        GarmentTest_Detail_Result Encode_Detail(string ID, string No, DetailStatus status);

        GarmentTest_Result SentMail(string ID, string No, List<Quality_MailGroup> mailGroups);

        GarmentTest_Detail_Result ToReport(string ID, string No, ReportType type, bool IsToPDF, bool test = false);

    }
}
