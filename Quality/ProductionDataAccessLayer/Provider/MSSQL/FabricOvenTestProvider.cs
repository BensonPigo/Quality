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
    public class FabricOvenTestProvider : SQLDAL, IFabricOvenTestProvider
    {
        #region 底層連線
        public FabricOvenTestProvider(string ConString) : base(ConString) { }
        public FabricOvenTestProvider(SQLDataTransaction tra) : base(tra) { }

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
        o.ReportNo,
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
from Oven o with (nolock)
LEFT JOIN SciPMSFile_Oven oi with (nolock) ON o.ID = oi.ID
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

        public FabricOvenTest_Result GetFabricOvenTest_Main(string POID)
        {
            FabricOvenTest_Result fabricOvenTest_Result = new FabricOvenTest_Result();

            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", POID);

            string sqlGetFabricOvenTest_Main = @"
declare @MtlLeadTime tinyint

select @MtlLeadTime = isnull(MtlLeadTime, 0)
from system WITH(NOLOCK)


select	[POID] = p.ID,
		[StyleID] = p.StyleID,
		[BrandID] = p.BrandID,
		[SeasonID] = p.SeasonId,
		[CutInline] = o.CutInline,
		[MinSciDelivery] = p.MinSciDelivery,
		[TargetLeadTime] = case when p.MinSciDelivery is null then null
								when o.CutInline < dateadd(Day, @MtlLeadTime, p.MinSciDelivery) then o.CutInline
								else dateadd(Day, @MtlLeadTime, p.MinSciDelivery) end,
		[CompletionDate] = iif(p.LabOvenPercent >= 100, (select max(InspDate) from oven WITH(NOLOCK) where POID = p.ID), null),
		[LabOvenPercent] = p.LabOvenPercent,
		[Remark] = p.OvenLaboratoryRemark,
		[CreateBy] = Concat(p.AddName, '-', pass1AddName.Name, ' ', Format(p.AddDate, 'yyyy/MM/dd HH:mm:ss')),
		[EditBy] = Concat(p.EditName, '-', pass1EditName.Name, ' ', Format(p.EditDate, 'yyyy/MM/dd HH:mm:ss'))
from PO p with (nolock)
inner join Orders o with (nolock) on p.ID = o.ID
left join pass1 pass1AddName WITH(NOLOCK) on p.AddName = pass1AddName.ID
left join pass1 pass1EditName WITH(NOLOCK) on p.EditName = pass1EditName.ID
where p.id = @POID
";

            IList<FabricOvenTest_Main> listFabricOvenTest_Main = ExecuteList<FabricOvenTest_Main>(CommandType.Text, sqlGetFabricOvenTest_Main, listPar);

            if (listFabricOvenTest_Main.Count == 0)
            {
                throw new Exception($"PO<{POID}> data not found");
            }

            fabricOvenTest_Result.Main = listFabricOvenTest_Main[0];

            string sqlGetDetails = @"
select	[TestNo] = cast(o.TestNo as varchar),
		o.ReportNo,
		[InspDate] = o.InspDate,
		[Article] = o.Article,
		[Result] = o.Result,
		[Inspector] = o.Inspector,
		[Remark] = o.Remark,
		[LastUpdate] = iif(	o.EditDate is null, 
							Concat(pass1AddName.Name, ' ', Format(o.AddDate, 'yyyy/MM/dd HH:mm:ss')),
							Concat(pass1EditName.Name, ' ', Format(o.EditDate, 'yyyy/MM/dd HH:mm:ss'))),
		[Status] = o.Status
from Oven o with (nolock)
left join pass1 pass1AddName WITH(NOLOCK) on o.AddName = pass1AddName.ID
left join pass1 pass1EditName WITH(NOLOCK) on o.EditName = pass1EditName.ID
where o.POID = @POID
";

            fabricOvenTest_Result.Details = ExecuteList<FabricOvenTest_Detail>(CommandType.Text, sqlGetDetails, listPar).ToList();

            return fabricOvenTest_Result;
        }

        public void EditFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", fabricOvenTest_Detail_Result.Main.POID);
            listPar.Add("@TestNo", fabricOvenTest_Detail_Result.Main.TestNo);
            listPar.Add("@InspDate", fabricOvenTest_Detail_Result.Main.InspDate);
            listPar.Add("@Article", fabricOvenTest_Detail_Result.Main.Article);
            listPar.Add("@Inspector", fabricOvenTest_Detail_Result.Main.Inspector);            
            listPar.Add("@Remark", fabricOvenTest_Detail_Result.Main.Remark ?? "");
            listPar.Add("@editName", userID);
            listPar.Add("@TestBeforePicture", fabricOvenTest_Detail_Result.Main.TestBeforePicture);
            listPar.Add("@TestAfterPicture", fabricOvenTest_Detail_Result.Main.TestAfterPicture);

            string sqlUpdateOven = @"
SET XACT_ABORT ON
-----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
update  Oven set    InspDate = @InspDate,
                    Article = @Article,
                    Inspector = @Inspector,
                    Remark = @Remark,
                    EditName = @editName,
                    EditDate = getdate()
where   POID = @POID and TestNo = @TestNo

update  SciPMSFile_Oven set  
                    TestBeforePicture = @TestBeforePicture,
                    TestAfterPicture = @TestAfterPicture
where   POID = @POID and TestNo = @TestNo


select  [OvenID] = ID
from    Oven WITH(NOLOCK)
where   POID = @POID and TestNo = @TestNo
";


            FabricOvenTest_Detail_Result oldOvenData = GetFabricOvenTest_Detail(fabricOvenTest_Detail_Result.Main.POID, fabricOvenTest_Detail_Result.Main.TestNo);

            List<FabricOvenTest_Detail_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<FabricOvenTest_Detail_Detail>(
                    fabricOvenTest_Detail_Result.Details,
                    oldOvenData.Details,
                    "OvenGroup,SEQ",
                    "Result,SubmitDate,Roll,ChangeScale,ResultChange,StainingScale,ResultStain,Remark,Temperature,Time");



            string sqlInsertOvenDetail = @"
insert into Oven_Detail(
ID             ,
OvenGroup      ,
SEQ1           ,
SEQ2           ,
Roll           ,
Dyelot         ,
Result         ,
changeScale    ,
StainingScale  ,
Remark         ,
AddName        ,
AddDate        ,
ResultChange   ,
ResultStain    ,
SubmitDate     ,
Temperature    ,
Time
)
values
(
@ID             ,
@OvenGroup      ,
@SEQ1           ,
@SEQ2           ,
@Roll           ,
@Dyelot         ,
@Result         ,
@changeScale    ,
@StainingScale  ,
@Remark         ,
@AddName        ,
getdate()        ,
@ResultChange   ,
@ResultStain    ,
@SubmitDate     ,
@Temperature    ,
@Time
)

";

            string sqlDeleteDetail = @"
delete  Oven_Detail
where   ID = @ID and
        OvenGroup = @OvenGroup and
        SEQ1 = @SEQ1 and
        SEQ2 = @SEQ2

exec UpdateInspPercent 'LabOven',@POID
";

            string sqlUpdateDetail = @"
update  Oven_Detail set Roll           =  @Roll         ,
                        Dyelot         =  @Dyelot       ,
                        Result         =  @Result       ,
                        changeScale    =  @changeScale  ,
                        StainingScale  =  @StainingScale,
                        Remark         =  @Remark       ,
                        ResultChange   =  @ResultChange ,
                        ResultStain    =  @ResultStain  ,
                        SubmitDate     =  @SubmitDate   ,
                        Temperature    =  @Temperature  ,
                        Time           =  @Time,
                        EditName = @EditName,
                        EditDate = getdate()
        where   ID = @ID and
                OvenGroup = @OvenGroup and
                SEQ1 = @SEQ1 and
                SEQ2 = @SEQ2
";

            using (TransactionScope transaction = new TransactionScope())
            {
                DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdateOven, listPar);
                long ovenID = (long)dtResult.Rows[0]["OvenID"];
                foreach (FabricOvenTest_Detail_Detail detailItem in needUpdateDetailList)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();
                    int temperature = 0;
                    int time = 0;
                    int.TryParse(detailItem.Temperature, out temperature);
                    int.TryParse(detailItem.Time, out time);
                    switch (detailItem.StateType)
                    {
                        case DatabaseObject.Public.CompareStateType.Add:
                            listDetailPar.Add("@ID", ovenID);
                            listDetailPar.Add("@OvenGroup", detailItem.OvenGroup);
                            listDetailPar.Add("@SEQ1", detailItem.Seq1);
                            listDetailPar.Add("@SEQ2", detailItem.Seq2);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@changeScale", detailItem.ChangeScale);
                            listDetailPar.Add("@StainingScale", detailItem.StainingScale);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@AddName", userID);
                            listDetailPar.Add("@ResultChange", detailItem.ResultChange);
                            listDetailPar.Add("@ResultStain", detailItem.ResultStain);
                            listDetailPar.Add("@SubmitDate", detailItem.SubmitDate);
                            listDetailPar.Add("@Temperature", temperature);
                            listDetailPar.Add("@Time", time);

                            ExecuteNonQuery(CommandType.Text, sqlInsertOvenDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Edit:
                            listDetailPar.Add("@ID", ovenID);
                            listDetailPar.Add("@OvenGroup", detailItem.OvenGroup);
                            listDetailPar.Add("@SEQ1", detailItem.Seq1);
                            listDetailPar.Add("@SEQ2", detailItem.Seq2);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@changeScale", detailItem.ChangeScale);
                            listDetailPar.Add("@StainingScale", detailItem.StainingScale);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@EditName", userID);
                            listDetailPar.Add("@ResultChange", detailItem.ResultChange);
                            listDetailPar.Add("@ResultStain", detailItem.ResultStain);
                            listDetailPar.Add("@SubmitDate", detailItem.SubmitDate);
                            listDetailPar.Add("@Temperature", temperature);
                            listDetailPar.Add("@Time", time);

                            ExecuteNonQuery(CommandType.Text, sqlUpdateDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Delete:
                            listDetailPar.Add("@ID", ovenID);
                            listDetailPar.Add("@OvenGroup", detailItem.OvenGroup);
                            listDetailPar.Add("@SEQ1", detailItem.Seq1);
                            listDetailPar.Add("@SEQ2", detailItem.Seq2);
                            listDetailPar.Add("@POID", fabricOvenTest_Detail_Result.Main.POID);

                            ExecuteNonQuery(CommandType.Text, sqlDeleteDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.None:
                            break;
                        default:
                            break;
                    }
                }

                string UpdateInspPercent = "exec UpdateInspPercent 'LabOven',@POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }

        public void AddFabricOvenTestDetail(FabricOvenTest_Detail_Result fabricOvenTest_Detail_Result, string userID, out string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", fabricOvenTest_Detail_Result.Main.POID);
            listPar.Add("@InspDate", fabricOvenTest_Detail_Result.Main.InspDate);
            listPar.Add("@Article", fabricOvenTest_Detail_Result.Main.Article);
            listPar.Add("@Inspector", fabricOvenTest_Detail_Result.Main.Inspector);
            listPar.Add("@Remark", fabricOvenTest_Detail_Result.Main.Remark ?? "");
            listPar.Add("@addName", userID);
            listPar.Add("@TestBeforePicture", fabricOvenTest_Detail_Result.Main.TestBeforePicture);
            listPar.Add("@TestAfterPicture", fabricOvenTest_Detail_Result.Main.TestAfterPicture);

            string NewReportNo = GetID(fabricOvenTest_Detail_Result.MDivisionID + "FO", "Oven", DateTime.Today, 2, "ReportNo");
            listPar.Add("@ReportNo", NewReportNo);

            string sqlInsertOven = @"
SET XACT_ABORT ON

declare @TestNo numeric(2,0)
DECLARE @OvenID table (ID bigint, TestNo numeric(2, 0))

select  @TestNo = isnull(Max(TestNo), 0) + 1
from    Oven  WITH(NOLOCK)
where POID = @POID

----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
insert into Oven(POID, TestNo, InspDate, Article, Status, Inspector, Remark, addName, addDate ,ReportNo)
        OUTPUT INSERTED.ID, INSERTED.TestNo into @OvenID
        values(@POID, @TestNo, @InspDate, @Article, 'New', @Inspector, @Remark, @addName, getdate() ,@ReportNo)

select  [OvenID] = ID, TestNo
from @OvenID

insert into SciPMSFile_Oven(ID, POID, TestNo, TestBeforePicture, TestAfterPicture)
        values(
(select ID from @OvenID)
, @POID, @TestNo, @TestBeforePicture, @TestAfterPicture)

";


            string sqlInsertOvenDetail = @"
insert into Oven_Detail(
ID             ,
OvenGroup      ,
SEQ1           ,
SEQ2           ,
Roll           ,
Dyelot         ,
Result         ,
changeScale    ,
StainingScale  ,
Remark         ,
AddName        ,
AddDate        ,
ResultChange   ,
ResultStain    ,
SubmitDate     ,
Temperature    ,
Time
)
values
(
@ID             ,
@OvenGroup      ,
@SEQ1           ,
@SEQ2           ,
@Roll           ,
@Dyelot         ,
@Result         ,
@changeScale    ,
@StainingScale  ,
@Remark         ,
@AddName        ,
getdate()        ,
@ResultChange   ,
@ResultStain    ,
@SubmitDate     ,
@Temperature    ,
@Time
)
";

            using (TransactionScope transaction = new TransactionScope())
            {
                DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlInsertOven, listPar);
                long ovenID = (long)dtResult.Rows[0]["OvenID"];
                TestNo = dtResult.Rows[0]["TestNo"].ToString();
                foreach (FabricOvenTest_Detail_Detail detailItem in fabricOvenTest_Detail_Result.Details)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();
                    listDetailPar.Add("@ID", ovenID);
                    listDetailPar.Add("@OvenGroup", detailItem.OvenGroup);
                    listDetailPar.Add("@SEQ1", detailItem.Seq1);
                    listDetailPar.Add("@SEQ2", detailItem.Seq2);
                    listDetailPar.Add("@Roll", detailItem.Roll);
                    listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                    listDetailPar.Add("@Result", detailItem.Result);
                    listDetailPar.Add("@changeScale", detailItem.ChangeScale);
                    listDetailPar.Add("@StainingScale", detailItem.StainingScale);
                    listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                    listDetailPar.Add("@AddName", userID);
                    listDetailPar.Add("@ResultChange", detailItem.ResultChange);
                    listDetailPar.Add("@ResultStain", detailItem.ResultStain);
                    listDetailPar.Add("@SubmitDate", detailItem.SubmitDate);
                    listDetailPar.Add("@Temperature", int.Parse(detailItem.Temperature));
                    listDetailPar.Add("@Time", int.Parse(detailItem.Time));

                    ExecuteNonQuery(CommandType.Text, sqlInsertOvenDetail, listDetailPar);
                }

                string UpdateInspPercent = "exec UpdateInspPercent 'LabOven',@POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }

        public void SaveFabricOvenTestMain(FabricOvenTest_Main fabricOvenTest_Main)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@Remark", fabricOvenTest_Main.Remark ?? "");
            listPar.Add("@POID", fabricOvenTest_Main.POID);

            string sqlUpdateOvenMain = @"
update PO set OvenLaboratoryRemark = @Remark
where ID = @POID

exec UpdateInspPercent 'LabOven',@POID
";
            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateOvenMain, listPar);
                transaction.Complete();
            }

        }

        public void EncodeFabricOven(string poID, string TestNo, string result)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);
            listPar.Add("@result", result);

            string sqlUpdateOvenMain = @"
update Oven set Status = 'Confirmed',
                Result = @result
where POID = @poID and TestNo = @TestNo

exec UpdateInspPercent 'LabOven',@poID
";
            ExecuteNonQuery(CommandType.Text, sqlUpdateOvenMain, listPar);
        }

        public void AmendFabricOven(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlUpdateOvenMain = @"
update Oven set Status = 'New',
                Result = ''
where POID = @poID and TestNo = @TestNo

exec UpdateInspPercent 'LabOven',@poID
";
            ExecuteNonQuery(CommandType.Text, sqlUpdateOvenMain, listPar);
        }

        public DataTable GetFailMailContentData(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlGetData = @"
select  [SP#] = ov.POID,
        [Style] = o.StyleID,
        [Brand] = o.BrandID,
        [Season] = o.SeasonID,
        [No of Test] = ov.TestNo,
        [Test Date] = Format(ov.InspDate, 'yyyy/MM/dd'),
        [Article] = ov.Article,
        [Result] = ov.Result,
        [Inspector] = ov.Inspector,
        [Remark] = ov.Remark,
        oi.TestBeforePicture,
        oi.TestAfterPicture
from Oven ov with (nolock)
left join Orders o with (nolock) on ov.POID = o.ID
left join SciPMSFile_Oven oi with (nolock) on oi.ID=ov.ID
where ov.POID = @poID and ov.TestNo = @TestNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public DataTable GetOvenDetailForExcel(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlGetData = @"
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
        [Time] = cast(od.Time as varchar),
        [Supplier] = ps.SuppID+'-'+s.AbbEN
from Oven_Detail od with (nolock)
inner join Oven o with (nolock) on o.ID = od.ID
left join PO_Supp_Detail psd with (nolock) on o.POID = psd.ID and od.SEQ1 = psd.SEQ1 and od.SEQ2 = psd.SEQ2
left join PO_Supp ps WITH (NOLOCK) on psd.ID = ps.ID and psd.Seq1 = ps.Seq1
left join supp s with (nolock) on ps.SuppID = s.ID
left join pass1 pass1EditName WITH(NOLOCK) on od.EditName = pass1EditName.ID
where   o.POID = @POID and o.TestNo = @TestNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public DataTable GetOven(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlGetData = @"
select  ov.ID
        ,ov.POID
        ,ov.TestNo
        ,ov.InspDate
        ,ov.Article
        ,ov.Result
        ,ov.Status
        ,ov.Inspector
        ,ov.Remark
        ,ov.addName
        ,ov.addDate
        ,ov.EditName
        ,ov.EditDate
        ,ov.Temperature
        ,ov.Time
        ,oi.TestBeforePicture
        ,oi.TestAfterPicture
        ,[InspectorName] = (select Name from Pass1 WITH(NOLOCK) where ID = ov.Inspector)
        ,ov.ReportNo
from    Oven ov with (nolock)
left join SciPMSFile_Oven oi with (nolock) on oi.ID=ov.ID
where   ov.POID = @poID and ov.TestNo = @TestNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public void DeleteOven(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlDeleteOven = @"
SET XACT_ABORT ON
delete  Oven_Detail where ID = (select ID from Oven where POID = @poID and TestNo = @TestNo)
delete  Oven where POID = @poID and TestNo = @TestNo
delete SciPMSFile_Oven where POID = @poID and TestNo = @TestNo
exec UpdateInspPercent 'LabOven',@poID
";
            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlDeleteOven, listPar);
                transaction.Complete();
            }
        }

        #endregion

    }
}
