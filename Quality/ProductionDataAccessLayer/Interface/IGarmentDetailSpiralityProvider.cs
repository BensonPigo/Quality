using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IGarmentDetailSpiralityProvider
    {
        IList<Garment_Detail_Spirality> Get_Garment_Detail_Spirality(Int64 ID, string No);

        bool Update_Spirality(List<Garment_Detail_Spirality> source);
    }
}
