using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using static MICS.DataAccessLayer.Provider.MSSQL.ColorFastnessDetailProvider;

namespace MICS.DataAccessLayer.Interface
{  
    public interface IColorFastnessDetailProvider
    {
        IList<Fabirc_ColorFastness_Detail_ViewModel> Get_DetailBody(string ID);

        IList<PO_Supp_Detail> Get_Seq(string POID, string Seq1, string Seq2);

        IList<FtyInventory> Get_Roll(string POID, string Seq1, string Seq2);
    }
}
