using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using System.Collections.Generic;
using System.Web.Mvc;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestProvider
    {
        IList<Style> GetStyleID();

        IList<Brand> GetBrandID();

        IList<Season> GetSeasonID();
        string GetFactoryNameEN(string factory);

        IList<GarmentTest> GetArticle(GarmentTest_ViewModel filter);

        IList<GarmentTest_ViewModel> Get_GarmentTest(GarmentTest_Request filter);

        IList<GarmentTest_ViewModel> Get(string ID);

        void Save_GarmentTest(GarmentTest_ViewModel master, List<GarmentTest_Detail_ViewModel> detail, string UserID, bool sameInstance);

        bool Update_GarmentTest_Result(string ID);
        void Save_New_FGPT_Item(GarmentTest_Detail_FGPT_ViewModel newItem);
        void Save_New_Shrinkage_Item(GarmentTest_Detail_Shrinkage newItem);
        
        void Delete_Original_FGPT_Item(GarmentTest_Detail_FGPT_ViewModel newItem);
        string CheckInstance();

        IList<GarmentTest_Detail_ViewModel> GetDetail_LastTestNo(GarmentTest_Request filter, string Type);
        List<SelectListItem> GetShrinkageLocation(long GarmentTestID);
    }
}
