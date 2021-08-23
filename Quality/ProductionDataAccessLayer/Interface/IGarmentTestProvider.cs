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
    }
}
