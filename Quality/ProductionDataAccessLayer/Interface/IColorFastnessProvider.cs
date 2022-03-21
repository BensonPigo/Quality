using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Data;

namespace MICS.DataAccessLayer.Interface
{
    public interface IColorFastnessProvider
    {
        List<string> GetScales();

        FabricColorFastness_ViewModel GetMain(string PoID);

        DateTime? Get_Target_LeadTime(object CUTINLINE, object MinSciDelivery);

        IList<ColorFastness_Result> Get(string ID);

        bool Save_PO(string PoID, string Remark);

        bool Delete_ColorFastness(string ID);

        bool Encode_ColorFastness(string ID, string Status, string Result, string UserID);

        DataTable Get_Mail_Content(string POID, string ID, string TestNo);

        string Get_InspectName(string Inspector);

        DataTable Get_PO_DataTable(string PoID);

        string Get_Supplier(string PoID, string Seq1);
    }
}
