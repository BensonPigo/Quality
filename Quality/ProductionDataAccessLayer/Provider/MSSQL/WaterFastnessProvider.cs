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

        public WaterFastness_Detail_Result GetWaterFastness_Detail(string poID, string TestNo ,string BrandID = "")
        {
            WaterFastness_Detail_Result WaterFastness_Detail_Result = new WaterFastness_Detail_Result();

            if (string.IsNullOrEmpty(TestNo))
            {
                WaterFastness_Detail_Result.Main.Status = "";
                WaterFastness_Detail_Result.Main.POID = poID;
                WaterFastness_Detail_Result.Main.BrandID = BrandID;
                return WaterFastness_Detail_Result;
            }

            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", poID);
            listPar.Add("@TestNo", decimal.Parse(TestNo));

            string sqlGetWaterFastness_Detail = @"

select	[TestNo] = cast(o.TestNo as varchar),
        [POID] = o.POID,
        o.ID, 
        o.ReportNo,
        od.BrandID,
		[InspDate] = o.InspDate,
		[Article] = o.Article,
		[Inspector] = o.Inspector,
        [InspectorName] = pass1Inspector.Name,
        [Result] = o.Result,
		[Remark] = o.Remark,
		[Status] = o.Status,
        Temperature =  o.Temperature ,
        Time =   o.Time ,
        [TestBeforePicture] = (select top 1 TestBeforePicture from SciPMSFile_WaterFastness oi WITH(NOLOCK) where o.ID = oi.ID ),
        [TestAfterPicture] = (select top 1 TestAfterPicture from SciPMSFile_WaterFastness oi WITH(NOLOCK) where o.ID = oi.ID )
from WaterFastness o with (nolock)
inner join Orders od with (nolock) on od.ID = o.POID
left join pass1 pass1Inspector WITH(NOLOCK) on o.Inspector = pass1Inspector.ID
where o.POID = @POID and o.TestNo = @TestNo
";

            IList<WaterFastness_Detail_Main> listWaterFastness_Detail = ExecuteList<WaterFastness_Detail_Main>(CommandType.Text, sqlGetWaterFastness_Detail, listPar);

            if (listWaterFastness_Detail.Count == 0)
            {
                throw new Exception($"TestNo<{TestNo}> data not found");
            }

            WaterFastness_Detail_Result.Main = listWaterFastness_Detail[0];

            string sqlGetDetails = @"
select	[SubmitDate] = od.SubmitDate,
        [WaterFastnessGroup] = od.WaterFastnessGroup,
        [SEQ] = Concat(od.Seq1, '-', od.Seq2),
        [Roll] = od.Roll,
        [Dyelot] = od.Dyelot,
        [Refno] = psd.Refno,
        [SCIRefno] = psd.SCIRefno,
        [ColorID] = pc.SpecValue,
        [Result] = od.Result
        ,od.ChangeScale
        ,od.AcetateScale
        ,od.CottonScale
        ,od.NylonScale
        ,od.PolyesterScale
        ,od.AcrylicScale
        ,od.WoolScale
        ,od.ResultChange
        ,od.ResultAcetate
        ,od.ResultCotton
        ,od.ResultNylon
        ,od.ResultPolyester
        ,od.ResultAcrylic
        ,od.ResultWool,
        [Remark] = od.Remark,
        [LastUpdate] = Concat(od.EditName, '-', pass1EditName.Name, ' ', pass1EditName.Extno),
        [Temperature] = o.Temperature ,
        [Time] = o.Time 
from WaterFastness_Detail od with (nolock)
inner join WaterFastness o with (nolock) on o.ID = od.ID
left join PO_Supp_Detail psd with (nolock) on o.POID = psd.ID and od.SEQ1 = psd.SEQ1 and od.SEQ2 = psd.SEQ2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join pass1 pass1EditName on od.EditName = pass1EditName.ID
where   o.POID = @POID and o.TestNo = @TestNo
";

            WaterFastness_Detail_Result.Details = ExecuteList<WaterFastness_Detail_Detail>(CommandType.Text, sqlGetDetails, listPar).ToList();

            return WaterFastness_Detail_Result;
        }

        public WaterFastness_Result GetWaterFastness_Main(string POID)
        {
            WaterFastness_Result WaterFastness_Result = new WaterFastness_Result();

            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", POID);

            string sqlGetWaterFastness_Main = @"
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
		[CompletionDate] = iif(p.LabWaterFastnessPercent >= 100, (select max(InspDate) from WaterFastness WITH(NOLOCK) where POID = p.ID), null),
		[LabWaterFastnessPercent] = p.LabWaterFastnessPercent,
		[Remark] = p.WaterFastnessLaboratoryRemark,
		[CreateBy] = Concat(p.AddName, '-', pass1AddName.Name, ' ', Format(p.AddDate, 'yyyy/MM/dd HH:mm:ss')),
		[EditBy] = Concat(p.EditName, '-', pass1EditName.Name, ' ', Format(p.EditDate, 'yyyy/MM/dd HH:mm:ss'))
from PO p with (nolock)
inner join Orders o with (nolock) on p.ID = o.ID
left join pass1 pass1AddName WITH(NOLOCK) on p.AddName = pass1AddName.ID
left join pass1 pass1EditName WITH(NOLOCK) on p.EditName = pass1EditName.ID
where p.id = @POID
";

            IList<WaterFastness_Main> listWaterFastness_Main = ExecuteList<WaterFastness_Main>(CommandType.Text, sqlGetWaterFastness_Main, listPar);

            if (listWaterFastness_Main.Count == 0)
            {
                throw new Exception($"PO<{POID}> data not found");
            }

            WaterFastness_Result.Main = listWaterFastness_Main[0];

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
from WaterFastness o with (nolock)
left join pass1 pass1AddName WITH(NOLOCK) on o.AddName = pass1AddName.ID
left join pass1 pass1EditName WITH(NOLOCK) on o.EditName = pass1EditName.ID
where o.POID = @POID
";

            WaterFastness_Result.Details = ExecuteList<WaterFastness_Detail>(CommandType.Text, sqlGetDetails, listPar).ToList();

            return WaterFastness_Result;
        }

        public void EditWaterFastnessDetail(WaterFastness_Detail_Result WaterFastness_Detail_Result, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", WaterFastness_Detail_Result.Main.POID);
            listPar.Add("@TestNo", WaterFastness_Detail_Result.Main.TestNo);
            listPar.Add("@InspDate", WaterFastness_Detail_Result.Main.InspDate);
            listPar.Add("@Article", WaterFastness_Detail_Result.Main.Article);
            listPar.Add("@Inspector", WaterFastness_Detail_Result.Main.Inspector);
            listPar.Add("@Remark", WaterFastness_Detail_Result.Main.Remark ?? "");
            listPar.Add("@editName", userID);
            listPar.Add("@TestBeforePicture", WaterFastness_Detail_Result.Main.TestBeforePicture);
            listPar.Add("@TestAfterPicture", WaterFastness_Detail_Result.Main.TestAfterPicture);
            listPar.Add("@Temperature", DbType.Int32, WaterFastness_Detail_Result.Main.Temperature);
            listPar.Add("@Time", DbType.Int32, WaterFastness_Detail_Result.Main.Time);

            string sqlUpdateWaterFastness = @"
SET XACT_ABORT ON
-----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
update  WaterFastness set    InspDate = @InspDate,
                    Article = @Article,
                    Inspector = @Inspector,
                    Temperature = @Temperature,
                    Time = @Time,
                    Remark = @Remark,
                    EditName = @editName,
                    EditDate = getdate()
where   POID = @POID and TestNo = @TestNo


select  [WaterFastnessID] = ID
from    WaterFastness WITH(NOLOCK)
where   POID = @POID and TestNo = @TestNo

update  SciPMSFile_WaterFastness set  
                    TestBeforePicture = @TestBeforePicture,
                    TestAfterPicture = @TestAfterPicture
where  ID IN (
    select   ID
    from    WaterFastness WITH(NOLOCK)
    where   POID = @POID and TestNo = @TestNo
)

";


            WaterFastness_Detail_Result oldWaterFastnessData = GetWaterFastness_Detail(WaterFastness_Detail_Result.Main.POID, WaterFastness_Detail_Result.Main.TestNo);

            List<WaterFastness_Detail_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<WaterFastness_Detail_Detail>(
                    WaterFastness_Detail_Result.Details,
                    oldWaterFastnessData.Details,
                    "WaterFastnessGroup,SEQ",
                    "Result,SubmitDate,Roll,ChangeScale,ResultChange" +
                    ",AcetateScale,ResultAcetate,CottonScale,ResultCotton" +
                    ",NylonScale,ResultNylon,PolyesterScale,ResultPolyester" +
                    ",AcrylicScale,ResultAcrylic" +
                    ",WoolScale,ResultWool,Remark");



            string sqlInsertWaterFastnessDetail = @"
insert into WaterFastness_Detail(
ID             ,
WaterFastnessGroup      ,
SEQ1           ,
SEQ2           ,
Roll           ,
Dyelot         ,
Result         ,
ChangeScale    ,
ResultChange   

,AcetateScale
,ResultAcetate
,CottonScale
,ResultCotton
,NylonScale
,ResultNylon
,PolyesterScale
,ResultPolyester
,AcrylicScale
,ResultAcrylic
,WoolScale
,ResultWool,

Remark         ,
AddName        ,
AddDate        ,
SubmitDate     
)
values
(
@ID             ,
@WaterFastnessGroup      ,
@SEQ1           ,
@SEQ2           ,
@Roll           ,
@Dyelot         ,
@Result         ,
@ChangeScale    ,
@ResultChange   
,@AcetateScale
,@ResultAcetate
,@CottonScale
,@ResultCotton
,@NylonScale
,@ResultNylon
,@PolyesterScale
,@ResultPolyester
,@AcrylicScale
,@ResultAcrylic
,@WoolScale
,@ResultWool,
@Remark         ,
@AddName        ,
getdate()        ,
@SubmitDate     
)

";

            string sqlDeleteDetail = @"
delete  WaterFastness_Detail
where   ID = @ID and
        WaterFastnessGroup = @WaterFastnessGroup and
        SEQ1 = @SEQ1 and
        SEQ2 = @SEQ2

exec UpdateInspPercent 'LabWaterFastness',@POID
";

            string sqlUpdateDetail = @"
update  WaterFastness_Detail set Roll           =  @Roll         ,
                        Dyelot         =  @Dyelot       ,
                        Result         =  @Result       ,
                        ChangeScale    =  @ChangeScale  ,
                        ResultChange   =  @ResultChange ,

                        AcetateScale    =  @AcetateScale  ,
                        ResultAcetate   =  @ResultAcetate ,
                        CottonScale    =  @CottonScale  ,
                        ResultCotton   =  @ResultCotton ,
                        NylonScale    =  @NylonScale  ,
                        ResultNylon   =  @ResultNylon ,
                        PolyesterScale    =  @PolyesterScale  ,
                        ResultPolyester   =  @ResultPolyester ,
                        AcrylicScale    =  @AcrylicScale  ,
                        ResultAcrylic   =  @ResultAcrylic ,
                        WoolScale    =  @WoolScale  ,
                        ResultWool   =  @ResultWool ,

                        Remark         =  @Remark       ,
                        SubmitDate     =  @SubmitDate   ,
                        EditName = @EditName,
                        EditDate = getdate()
        where   ID = @ID and
                WaterFastnessGroup = @WaterFastnessGroup and
                SEQ1 = @SEQ1 and
                SEQ2 = @SEQ2
";

            using (TransactionScope transaction = new TransactionScope())
            {
                DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdateWaterFastness, listPar);
                long WaterFastnessID = (long)dtResult.Rows[0]["WaterFastnessID"];
                foreach (WaterFastness_Detail_Detail detailItem in needUpdateDetailList)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();

                    switch (detailItem.StateType)
                    {
                        case DatabaseObject.Public.CompareStateType.Add:
                            listDetailPar.Add("@ID", WaterFastnessID);
                            listDetailPar.Add("@WaterFastnessGroup", detailItem.WaterFastnessGroup);
                            listDetailPar.Add("@SEQ1", detailItem.Seq1);
                            listDetailPar.Add("@SEQ2", detailItem.Seq2);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Result", detailItem.Result);

                            listDetailPar.Add("@ChangeScale", detailItem.ChangeScale);
                            listDetailPar.Add("@ResultChange", detailItem.ResultChange);

                            listDetailPar.Add("@AcetateScale", detailItem.AcetateScale);
                            listDetailPar.Add("@ResultAcetate", detailItem.ResultAcetate);

                            listDetailPar.Add("@CottonScale", detailItem.CottonScale);
                            listDetailPar.Add("@ResultCotton", detailItem.ResultCotton);

                            listDetailPar.Add("@NylonScale", detailItem.NylonScale);
                            listDetailPar.Add("@ResultNylon", detailItem.ResultNylon);

                            listDetailPar.Add("@PolyesterScale", detailItem.PolyesterScale);
                            listDetailPar.Add("@ResultPolyester", detailItem.ResultPolyester);

                            listDetailPar.Add("@AcrylicScale", detailItem.AcrylicScale);
                            listDetailPar.Add("@ResultAcrylic", detailItem.ResultAcrylic);

                            listDetailPar.Add("@WoolScale", detailItem.WoolScale);
                            listDetailPar.Add("@ResultWool", detailItem.ResultWool);

                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@AddName", userID);
                            listDetailPar.Add("@SubmitDate", detailItem.SubmitDate);

                            ExecuteNonQuery(CommandType.Text, sqlInsertWaterFastnessDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Edit:
                            listDetailPar.Add("@ID", WaterFastnessID);
                            listDetailPar.Add("@WaterFastnessGroup", detailItem.WaterFastnessGroup);
                            listDetailPar.Add("@SEQ1", detailItem.Seq1);
                            listDetailPar.Add("@SEQ2", detailItem.Seq2);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@ChangeScale", detailItem.ChangeScale);
                            listDetailPar.Add("@ResultChange", detailItem.ResultChange);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@EditName", userID);


                            listDetailPar.Add("@ChangeScale", detailItem.ChangeScale);
                            listDetailPar.Add("@ResultChange", detailItem.ResultChange);

                            listDetailPar.Add("@AcetateScale", detailItem.AcetateScale);
                            listDetailPar.Add("@ResultAcetate", detailItem.ResultAcetate);

                            listDetailPar.Add("@CottonScale", detailItem.CottonScale);
                            listDetailPar.Add("@ResultCotton", detailItem.ResultCotton);

                            listDetailPar.Add("@NylonScale", detailItem.NylonScale);
                            listDetailPar.Add("@ResultNylon", detailItem.ResultNylon);

                            listDetailPar.Add("@PolyesterScale", detailItem.PolyesterScale);
                            listDetailPar.Add("@ResultPolyester", detailItem.ResultPolyester);

                            listDetailPar.Add("@AcrylicScale", detailItem.AcrylicScale);
                            listDetailPar.Add("@ResultAcrylic", detailItem.ResultAcrylic);

                            listDetailPar.Add("@WoolScale", detailItem.WoolScale);
                            listDetailPar.Add("@ResultWool", detailItem.ResultWool);

                            listDetailPar.Add("@SubmitDate", detailItem.SubmitDate);

                            ExecuteNonQuery(CommandType.Text, sqlUpdateDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Delete:
                            listDetailPar.Add("@ID", WaterFastnessID);
                            listDetailPar.Add("@WaterFastnessGroup", detailItem.WaterFastnessGroup);
                            listDetailPar.Add("@SEQ1", detailItem.Seq1);
                            listDetailPar.Add("@SEQ2", detailItem.Seq2);
                            listDetailPar.Add("@POID", WaterFastness_Detail_Result.Main.POID);

                            ExecuteNonQuery(CommandType.Text, sqlDeleteDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.None:
                            break;
                        default:
                            break;
                    }
                }

                string UpdateInspPercent = "exec UpdateInspPercent 'LabWaterFastness',@POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }

        public void AddWaterFastnessDetail(WaterFastness_Detail_Result waterFastness_Detail_Result, string userID, out string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", waterFastness_Detail_Result.Main.POID);
            listPar.Add("@InspDate", waterFastness_Detail_Result.Main.InspDate);
            listPar.Add("@Article", waterFastness_Detail_Result.Main.Article);
            listPar.Add("@Inspector", waterFastness_Detail_Result.Main.Inspector);
            listPar.Add("@Remark", waterFastness_Detail_Result.Main.Remark ?? "");
            listPar.Add("@addName", userID);
            listPar.Add("@TestBeforePicture", waterFastness_Detail_Result.Main.TestBeforePicture);
            listPar.Add("@TestAfterPicture", waterFastness_Detail_Result.Main.TestAfterPicture);
            listPar.Add("@Temperature", DbType.Int32, waterFastness_Detail_Result.Main.Temperature);
            listPar.Add("@Time", DbType.Int32, waterFastness_Detail_Result.Main.Time);

            string NewReportNo = GetID(waterFastness_Detail_Result.MDivisionID + "WF", "WaterFastness", DateTime.Today, 2, "ReportNo");
            listPar.Add("@ReportNo", NewReportNo);


            string sqlInsertWaterFastness = @"
SET XACT_ABORT ON

declare @TestNo numeric(2,0)
DECLARE @WaterFastnessID table (ID bigint, TestNo numeric(2, 0))

select  @TestNo = isnull(Max(TestNo), 0) + 1
from    WaterFastness  WITH(NOLOCK)
where POID = @POID

----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
insert into WaterFastness(POID, TestNo, InspDate, Article, Status, Inspector, Temperature, Time , Remark, addName, addDate ,ReportNo)
        OUTPUT INSERTED.ID, INSERTED.TestNo into @WaterFastnessID
        values(@POID, @TestNo, @InspDate, @Article, 'New', @Inspector, @Temperature, @Time, @Remark, @addName, getdate() ,@ReportNo)

select  [WaterFastnessID] = ID, TestNo
from @WaterFastnessID

insert into SciPMSFile_WaterFastness(ID, TestBeforePicture, TestAfterPicture)
        values(
(select ID from @WaterFastnessID) , @TestBeforePicture, @TestAfterPicture)
";

            string sqlInsertWaterFastnessDetail = @"
insert into WaterFastness_Detail(
ID             ,
WaterFastnessGroup      ,
SEQ1           ,
SEQ2           ,
Roll           ,
Dyelot         ,
Result         ,

ChangeScale    ,
ResultChange   ,
AcetateScale    ,
ResultAcetate   ,
CottonScale    ,
ResultCotton   ,
NylonScale    ,
ResultNylon   ,
PolyesterScale    ,
ResultPolyester   ,
AcrylicScale    ,
ResultAcrylic   ,
WoolScale    ,
ResultWool   ,

Remark         ,
AddName        ,
AddDate        ,
SubmitDate     
--Temperature    ,
--Time
)
values
(
@ID             ,
@WaterFastnessGroup      ,
@SEQ1           ,
@SEQ2           ,
@Roll           ,
@Dyelot         ,
@Result         ,

@ChangeScale    ,
@ResultChange   ,
@AcetateScale    ,
@ResultAcetate   ,
@CottonScale    ,
@ResultCotton   ,
@NylonScale    ,
@ResultNylon   ,
@PolyesterScale    ,
@ResultPolyester   ,
@AcrylicScale    ,
@ResultAcrylic   ,
@WoolScale    ,
@ResultWool   ,

@Remark         ,
@AddName        ,
getdate()        ,
@SubmitDate     
--@Temperature    ,
--@Time
)

";

            using (TransactionScope transaction = new TransactionScope())
            {
                DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlInsertWaterFastness, listPar);
                long WaterFastnessID = (long)dtResult.Rows[0]["WaterFastnessID"];
                TestNo = dtResult.Rows[0]["TestNo"].ToString();
                foreach (WaterFastness_Detail_Detail detailItem in waterFastness_Detail_Result.Details)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();
                    listDetailPar.Add("@ID", WaterFastnessID);
                    listDetailPar.Add("@WaterFastnessGroup", detailItem.WaterFastnessGroup);
                    listDetailPar.Add("@SEQ1", detailItem.Seq1);
                    listDetailPar.Add("@SEQ2", detailItem.Seq2);
                    listDetailPar.Add("@Roll", detailItem.Roll);
                    listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                    listDetailPar.Add("@Result", detailItem.Result);
                    listDetailPar.Add("@changeScale", detailItem.ChangeScale);
                    listDetailPar.Add("@ResultChange", detailItem.ResultChange);

                    listDetailPar.Add("@AcetateScale", detailItem.AcetateScale);
                    listDetailPar.Add("@ResultAcetate", detailItem.ResultAcetate);
                    listDetailPar.Add("@CottonScale", detailItem.CottonScale);
                    listDetailPar.Add("@ResultCotton", detailItem.ResultCotton);
                    listDetailPar.Add("@NylonScale", detailItem.NylonScale);
                    listDetailPar.Add("@ResultNylon", detailItem.ResultNylon);
                    listDetailPar.Add("@PolyesterScale", detailItem.PolyesterScale);
                    listDetailPar.Add("@ResultPolyester", detailItem.ResultPolyester);
                    listDetailPar.Add("@AcrylicScale", detailItem.AcrylicScale);
                    listDetailPar.Add("@ResultAcrylic", detailItem.ResultAcrylic);
                    listDetailPar.Add("@WoolScale", detailItem.WoolScale);
                    listDetailPar.Add("@ResultWool", detailItem.ResultWool);

                    listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                    listDetailPar.Add("@AddName", userID);
                    listDetailPar.Add("@SubmitDate", detailItem.SubmitDate);

                    ExecuteNonQuery(CommandType.Text, sqlInsertWaterFastnessDetail, listDetailPar);
                }

                string UpdateInspPercent = "exec UpdateInspPercent 'LabWaterFastness',@POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }

        public void SaveWaterFastnessMain(WaterFastness_Main WaterFastness_Main)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@Remark", WaterFastness_Main.Remark ?? "");
            listPar.Add("@POID", WaterFastness_Main.POID);

            string sqlUpdateWaterFastnessMain = @"
update PO set WaterFastnessLaboratoryRemark = @Remark
where ID = @POID

exec UpdateInspPercent 'LabWaterFastness',@POID
";
            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateWaterFastnessMain, listPar);
                transaction.Complete();
            }

        }

        public void EncodeWaterFastness(string poID, string TestNo, string result)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);
            listPar.Add("@result", result);

            string sqlUpdateWaterFastnessMain = @"
update WaterFastness set Status = 'Confirmed',
                Result = @result
where POID = @poID and TestNo = @TestNo

exec UpdateInspPercent 'LabWaterFastness',@poID
";
            ExecuteNonQuery(CommandType.Text, sqlUpdateWaterFastnessMain, listPar);
        }

        public void AmendWaterFastness(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlUpdateWaterFastnessMain = @"
update WaterFastness set Status = 'New',
                Result = ''
where POID = @poID and TestNo = @TestNo

exec UpdateInspPercent 'LabWaterFastness',@poID
";
            ExecuteNonQuery(CommandType.Text, sqlUpdateWaterFastnessMain, listPar);
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
        ov.ID
from    WaterFastness ov with (nolock)
left join  Orders o with (nolock) on ov.POID = o.ID
where ov.POID = @poID and ov.TestNo = @TestNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public DataTable GetWaterFastnessDetailForExcel(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlGetData = @"
select	[SubmitDate] = od.SubmitDate,
        [WaterFastnessGroup] = od.WaterFastnessGroup,
        [SEQ] = Concat(od.Seq1, '-', od.Seq2),
        [Roll] = od.Roll,
        [Dyelot] = od.Dyelot,
        [Refno] = psd.Refno,
        [SCIRefno] = psd.SCIRefno,
        [ColorID] = pc.SpecValue,
        [Result] = od.Result,
        [ChangeScale] = od.ChangeScale,
        [ResultChange] = od.ResultChange,

        [AcetateScale] = od.AcetateScale,
        [ResultAcetate] = od.ResultAcetate,
        [CottonScale] = od.CottonScale,
        [ResultCotton] = od.ResultCotton,
        [NylonScale] = od.NylonScale,
        [ResultNylon] = od.ResultNylon,
        [PolyesterScale] = od.PolyesterScale,
        [ResultPolyester] = od.ResultPolyester,
        [AcrylicScale] = od.AcrylicScale,
        [ResultAcrylic] = od.ResultAcrylic,
        [WoolScale] = od.WoolScale,
        [ResultWool] = od.ResultWool,

        [Remark] = od.Remark,
        [LastUpdate] = Concat(od.EditName, '-', pass1EditName.Name, ' ', pass1EditName.Extno),
        [Temperature] = cast(o.Temperature as varchar),
        [Time] = cast(o.Time as varchar),
        [Supplier] = ps.SuppID+'-'+s.AbbEN
from WaterFastness_Detail od with (nolock)
inner join WaterFastness o with (nolock) on o.ID = od.ID
left join PO_Supp_Detail psd with (nolock) on o.POID = psd.ID and od.SEQ1 = psd.SEQ1 and od.SEQ2 = psd.SEQ2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join PO_Supp ps WITH (NOLOCK) on psd.ID = ps.ID and psd.Seq1 = ps.Seq1
left join supp s with (nolock) on ps.SuppID = s.ID
left join pass1 pass1EditName WITH(NOLOCK) on od.EditName = pass1EditName.ID
where   o.POID = @POID and o.TestNo = @TestNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public DataTable GetWaterFastness(string poID, string TestNo)
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
        ,TestBeforePicture = (select top 1 TestBeforePicture from SciPMSFile_WaterFastness oi WITH(NOLOCK) where  oi.ID=ov.ID)
        ,TestAfterPicture = (select top 1 TestAfterPicture SciPMSFile_WaterFastness oi WITH(NOLOCK) where  oi.ID=ov.ID)
        ,[InspectorName] = (select Name from Pass1 WITH(NOLOCK) where ID = ov.Inspector)
from    WaterFastness ov with (nolock)
where   ov.POID = @poID and ov.TestNo = @TestNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public void DeleteWaterFastness(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlDeleteWaterFastness = @"
SET XACT_ABORT ON
delete  WaterFastness_Detail where ID = (select ID from WaterFastness where POID = @poID and TestNo = @TestNo)
delete  SciPMSFile_WaterFastness where ID = (select ID from WaterFastness where POID = @poID and TestNo = @TestNo)
delete  WaterFastness where POID = @poID and TestNo = @TestNo
exec UpdateInspPercent 'LabWaterFastness',@poID
";
            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlDeleteWaterFastness, listPar);
                transaction.Complete();
            }
        }

        public IList<WaterFastness_Excel> GetExcel(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            string sqlcmd = @"
select cd.SubmitDate
        ,o.SeasonID
        ,o.BrandID
        ,o.StyleID
        ,c.POID
        ,cd.Roll
        ,cd.Dyelot
        ,SCIRefno_Color = psd.SCIRefno + ' ' + pc.SpecValue
        ,c.Temperature
        ,c.Time
        ,cd.ChangeScale
        ,cd.AcetateScale
        ,cd.CottonScale
        ,cd.NylonScale
        ,cd.PolyesterScale
        ,cd.AcrylicScale
        ,cd.WoolScale
        ,cd.ResultChange
        ,cd.ResultAcetate
        ,cd.ResultCotton
        ,cd.ResultNylon
        ,cd.ResultPolyester
        ,cd.ResultAcrylic
        ,cd.ResultWool
        ,cd.Remark
        ,c.Inspector
        ,TestBeforePicture = (select top 1 TestBeforePicture from SciPMSFile_WaterFastness pmsFile WITH(NOLOCK) where pmsFile.ID =  cd.ID)
        ,TestAfterPicture = (select top 1 TestAfterPicture from SciPMSFile_WaterFastness pmsFile WITH(NOLOCK) where pmsFile.ID =  cd.ID)
        ,c.ReportNo
from WaterFastness_Detail cd WITH(NOLOCK)
left join WaterFastness c WITH(NOLOCK) on c.ID =  cd.ID
left join Orders o WITH(NOLOCK) on o.ID=c.POID
left join PO_Supp_Detail psd WITH(NOLOCK) on c.POID = psd.ID and cd.SEQ1 = psd.SEQ1 and cd.SEQ2 = psd.SEQ2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Pass1 pEdit WITH(NOLOCK) on pEdit.ID = cd.EditName
left join pass1 pAdd WITH(NOLOCK) on pAdd.ID = cd.AddName
where cd.ID = @ID
order by cd.SubmitDate
";
            var detail = ExecuteList<WaterFastness_Excel>(CommandType.Text, sqlcmd, objParameter);

            return detail.Any() ? detail : new List<WaterFastness_Excel>();
        }
    }
}
