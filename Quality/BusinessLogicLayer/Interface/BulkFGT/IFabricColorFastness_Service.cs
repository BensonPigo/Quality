using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BusinessLogicLayer.Service.BulkFGT.FabricColorFastness_Service;
using static MICS.DataAccessLayer.Provider.MSSQL.ColorFastnessDetailProvider;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IFabricColorFastness_Service
    {
        List<string> Get_Scales();

        FabricColorFastness_ViewModel Get_Main(string PoID);

        ColorFastness_Result GetDetailHeader(string ID);

        Fabric_ColorFastness_Detail_ViewModel GetDetailBody(string ID);

        IList<PO_Supp_Detail> GetSeq(string POID, string Seq1 = "", string Seq2 = "");

        IList<FtyInventory> GetRoll(string POID, string Seq1, string Seq2);

        BaseResult Save_ColorFastness_1stPage(string PoID, string Remark, List<ColorFastness_Result> _ColorFastness);

        BaseResult Save_ColorFastness_2ndPage(Fabric_ColorFastness_Detail_ViewModel source, string Mdivision, string UserID);

        Fabric_ColorFastness_Detail_ViewModel Encode_ColorFastness(Fabric_ColorFastness_Detail_ViewModel source, DetailStatus status, string UserID);

        BaseResult SentMail(string POID, string ID, List<Quality_MailGroup> mailGroups);

        Fabric_ColorFastness_Detail_ViewModel ToPDF(string ID, bool test);

        Fabric_ColorFastness_Detail_ViewModel ToExcel(string ID, bool test);
    }
}
