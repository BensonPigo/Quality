using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;

namespace MICS.DataAccessLayer.Interface
{
    public interface IColorFastnessProvider
    {
        List<string> GetScales();

        FabricColorFastness_ViewModel GetMain(string PoID);

        DateTime? Get_Target_LeadTime(object CUTINLINE, object MinSciDelivery);

        IList<ColorFastness_Result> Get(string ID);

        bool Save_PO(string PoID, string Remark);

        bool Delete_ColorFastness(string PoID, List<ColorFastness_Result> source);
    }
}
