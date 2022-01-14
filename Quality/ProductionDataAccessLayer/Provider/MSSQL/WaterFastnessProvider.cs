using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Data.SqlClient;
using DatabaseObject.ViewModel.FinalInspection;
using System.Linq;
using ToolKit;
using System.Web.Mvc;
using DatabaseObject.ResultModel;
using DatabaseObject;
using System.Transactions;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class WaterFastnessProvider : SQLDAL
    {
        #region 底層連線
        public WaterFastnessProvider(string ConString) : base(ConString) { }
        public WaterFastnessProvider(SQLDataTransaction tra) : base(tra) { }

        #endregion

        public FabricOvenTest_Detail_Result GetFabricOvenTest_Detail(string poID, string TestNo)
        {
            FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result = new FabricOvenTest_Detail_Result();

            if (string.IsNullOrEmpty(TestNo))
            {
                fabricOvenTest_Detail_Result.Main.Status = "";
                fabricOvenTest_Detail_Result.Main.POID = poID;
                return fabricOvenTest_Detail_Result;
            }

            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", poID);
            listPar.Add("@TestNo", decimal.Parse(TestNo));

            string sqlGetFabricOvenTest_Detail = @"

select	[TestNo] = cast(o.TestNo as varchar),
        [POID] = o.POID,
		[InspDate] = o.InspDate,
		[Article] = o.Article,
		[Inspector] = o.Inspector,
        [InspectorName] = pass1Inspector.Name,
        [Result] = o.Result,
		[Remark] = o.Remark,
		[Status] = o.Status,
        [TestBeforePicture] = oi.TestBeforePicture,
        [TestAfterPicture] = oi.TestAfterPicture
from WaterFastness o with (nolock)
LEFT JOIN [ExtendServer].PMSFile.dbo.WaterFastness oi with (nolock) ON o.ID = oi.ID
left join pass1 pass1Inspector WITH(NOLOCK) on o.Inspector = pass1Inspector.ID
where o.POID = @POID and o.TestNo = @TestNo
";

            IList<FabricOvenTest_Detail_Main> listFabricOvenTest_Detail = ExecuteList<FabricOvenTest_Detail_Main>(CommandType.Text, sqlGetFabricOvenTest_Detail, listPar);

            if (listFabricOvenTest_Detail.Count == 0)
            {
                throw new Exception($"TestNo<{TestNo}> data not found");
            }

            fabricOvenTest_Detail_Result.Main = listFabricOvenTest_Detail[0];

            string sqlGetDetails = @"
select	[SubmitDate] = od.SubmitDate,
        [OvenGroup] = od.OvenGroup,
        [SEQ] = Concat(od.Seq1, '-', od.Seq2),
        [Roll] = od.Roll,
        [Dyelot] = od.Dyelot,
        [Refno] = psd.Refno,
        [SCIRefno] = psd.SCIRefno,
        [ColorID] = psd.ColorID,
        [Result] = od.Result,
        [ChangeScale] = od.changeScale,
        [ResultChange] = od.ResultChange,
        [StainingScale] = od.StainingScale,
        [ResultStain] = od.ResultStain,
        [Remark] = od.Remark,
        [LastUpdate] = Concat(od.EditName, '-', pass1EditName.Name, ' ', pass1EditName.Extno),
        [Temperature] = cast(od.Temperature as varchar),
        [Time] = cast(od.Time as varchar)
from Oven_Detail od with (nolock)
inner join Oven o with (nolock) on o.ID = od.ID
left join PO_Supp_Detail psd with (nolock) on o.POID = psd.ID and od.SEQ1 = psd.SEQ1 and od.SEQ2 = psd.SEQ2
left join pass1 pass1EditName on od.EditName = pass1EditName.ID
where   o.POID = @POID and o.TestNo = @TestNo
";

            fabricOvenTest_Detail_Result.Details = ExecuteList<FabricOvenTest_Detail_Detail>(CommandType.Text, sqlGetDetails, listPar).ToList();

            return fabricOvenTest_Detail_Result;
        }
    }
}
