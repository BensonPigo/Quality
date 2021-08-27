using DatabaseObject.ProductionDB;
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

        IList<GarmentTest_ViewModel> Get_GarmentTest(GarmentTest_ViewModel filter);

        int Save_GarmentTest(GarmentTest_ViewModel master, List<GarmentTest_Detail> detail);
    }
}
