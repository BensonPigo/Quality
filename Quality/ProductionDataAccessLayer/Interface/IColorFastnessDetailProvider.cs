using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Data;
using static MICS.DataAccessLayer.Provider.MSSQL.ColorFastnessDetailProvider;

namespace MICS.DataAccessLayer.Interface
{  
    public interface IColorFastnessDetailProvider
    {
        Fabric_ColorFastness_Detail_ViewModel Get_DetailBody(string ID);

        IList<PO_Supp_Detail> Get_Seq(string POID, string Seq1, string Seq2);

        IList<FtyInventory> Get_Roll(string POID, string Seq1, string Seq2);

        bool Save_ColorFastness(Fabric_ColorFastness_Detail_ViewModel source, string Mdivision, string UserID, out string NewID);

        bool Delete_ColorFastness_Detail(string ID, List<Fabric_ColorFastness_Detail_Result> source);

        DataTable Get_SubmitDate(string ID);
        IList<ColorFastness_Excel> GetExcel(string ID);
    }
}
