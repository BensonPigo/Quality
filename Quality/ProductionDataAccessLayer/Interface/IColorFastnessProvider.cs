using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;

namespace MICS.DataAccessLayer.Interface
{
    public interface IColorFastnessProvider
    {
        List<string> GetScales();

        IList<FabricColorFastness_ViewModel> GetMain(string PoID);

        DateTime? Get_Target_LeadTime(object CUTINLINE, object MinSciDelivery);

        IList<ColorFastness> Get(string ID);
    }
}
