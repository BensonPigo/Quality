using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static MICS.DataAccessLayer.Provider.MSSQL.ColorFastnessDetailProvider;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IFabricColorFastness_Service
    {
        List<string> Get_Scales();

        FabricColorFastness_ViewModel Get_Main(string PoID);

        FabricColorFastness_ViewModel GetDetailHeader(string ID);

        IList<Fabirc_ColorFastness_Detail_ViewModel> GetDetailBody(string ID);

        IList<PO_Supp_Detail> GetSeq(string POID, string Seq1 = "", string Seq2 = "");

        IList<FtyInventory> GetRoll(string POID, string Seq1, string Seq2);
    }
}
