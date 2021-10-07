using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentTestProvider
    {
        IList<Style> GetStyleID();

        IList<Brand> GetBrandID();

        IList<Season> GetSeasonID();

        IList<GarmentTest> GetArticle(GarmentTest_ViewModel filter);

        IList<GarmentTest_ViewModel> Get_GarmentTest(GarmentTest_Request filter);

        IList<GarmentTest_ViewModel> Get(string ID);

        void Save_GarmentTest(GarmentTest_ViewModel master, List<GarmentTest_Detail_ViewModel> detail, string UserID);

        bool Update_GarmentTest_Result(string ID);
    }
}
