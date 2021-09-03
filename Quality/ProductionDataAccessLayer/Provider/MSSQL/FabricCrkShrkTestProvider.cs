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
    public class FabricCrkShrkTestProvider : SQLDAL, IFabricCrkShrkTestProvider
    {
        #region 底層連線
        public FabricCrkShrkTestProvider(string ConString) : base(ConString) { }
        public FabricCrkShrkTestProvider(SQLDataTransaction tra) : base(tra) { }

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
        [Result] = o.Result,
		[Remark] = o.Remark,
		[Status] = o.Status,
        [TestBeforePicture] = o.TestBeforePicture,
        [TestAfterPicture] = o.TestAfterPicture
from Oven o with (nolock)
left join pass1 pass1AddName on o.AddName = pass1AddName.ID
left join pass1 pass1EditName on o.EditName = pass1EditName.ID
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

        public FabricCrkShrkTest_Result GetFabricCrkShrkTest_Main(string POID)
        {
            FabricCrkShrkTest_Result fabricCrkShrkTest_Result = new FabricCrkShrkTest_Result();

            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", POID);

            string sqlGetFabricOvenTest_Main = @"
declare @MtlLeadTime tinyint

select @MtlLeadTime = isnull(MtlLeadTime, 0)
from system


select	[POID] = p.ID,
		[StyleID] = p.StyleID,
		[BrandID] = p.BrandID,
		[SeasonID] = p.SeasonId,
		[CutInline] = o.CutInline,
		[MinSciDelivery] = p.MinSciDelivery,
		[TargetLeadTime] = case when p.MinSciDelivery is null then null
								when o.CutInline < dateadd(Day, @MtlLeadTime, p.MinSciDelivery) then o.CutInline
								else dateadd(Day, @MtlLeadTime, p.MinSciDelivery) end,
		[CompletionDate] = iif(p.LabOvenPercent >= 100, FIR_CompletionDate.val, null),
		[FIRLabInspPercent] = p.FIRLabInspPercent,
        [complete] = p.complete,
		[Remark] = p.FirLaboratoryRemark,
		[CreateBy] = Concat(p.AddName, '-', pass1AddName.Name, ' ', Format(p.AddDate, 'yyyy/MM/dd HH:mm:ss')),
		[EditBy] = Concat(p.EditName, '-', pass1EditName.Name, ' ', Format(p.EditDate, 'yyyy/MM/dd HH:mm:ss'))
from PO p with (nolock)
inner join Orders o with (nolock) on p.ID = o.ID
left join pass1 pass1AddName on p.AddName = pass1AddName.ID
left join pass1 pass1EditName on p.EditName = pass1EditName.ID
outer apply (    select [val] = max(CompletionDate)
                from    (
                 select [CompletionDate] = Max(CrockingDate) from FIR_Laboratory where POID = p.ID
                 union all
                 select [CompletionDate] = Max(HeatDate) from FIR_Laboratory where POID = p.ID
                 union all
                 select [CompletionDate] = Max(WashDate) from FIR_Laboratory where POID = p.ID
                ) a
            ) FIR_CompletionDate
where p.id = @POID
";

            IList<FabricCrkShrkTest_Main> listFabricCrkShrkTest_Main = ExecuteList<FabricCrkShrkTest_Main>(CommandType.Text, sqlGetFabricOvenTest_Main, listPar);

            if (listFabricCrkShrkTest_Main.Count == 0)
            {
                throw new Exception($"PO<{POID}> data not found");
            }

            fabricCrkShrkTest_Result.Main = listFabricCrkShrkTest_Main[0];

            string sqlGetDetails = @"
select	[ID] = f.ID,
        [Seq] = Concat(f.Seq1, ' ', f.Seq2),
        [WKNo] = r.ExportID,
        [WhseArrival] = r.WhseArrival,
        [SCIRefno] = f.SCIRefno,
        [Refno] = f.Refno,
        [ColorID] = psd.ColorID,
        [Supplier] = Concat(f.SuppID, s.AbbEn),
        [ArriveQty] = f.ArriveQty,
        [ReceiveSampleDate] = fl.ReceiveSampleDate,
        [InspDeadline] = fl.InspDeadline,
        [AllResult] = fl.Result,
        [NonCrocking] = fl.NonCrocking,
        [Crocking] = fl.Crocking,
        [CrockingDate] = fl.CrockingDate,
        [CrockingRemark] = fl.CrockingRemark,
        [NonHeat] = fl.nonHeat,
        [Heat] = fl.Heat,
        [HeatDate] = fl.HeatDate,
        [HeatRemark] = fl.HeatRemark,
        [NonWash] = fl.nonWash,
        [Wash] = fl.Wash,
        [WashDate] = fl.WashDate,
        [WashRemark] = fl.WashRemark,
        [ReceivingID] = f.ReceivingID
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join Supp s with (nolock) on s.ID = f.SuppID
where f.POID = @POID
";

            fabricCrkShrkTest_Result.Details = ExecuteList<FabricCrkShrkTest_Detail>(CommandType.Text, sqlGetDetails, listPar).ToList();

            return fabricCrkShrkTest_Result;
        }


        public void SaveFabricCrkShrkTest_Main(FabricCrkShrkTest_Result fabricCrkShrkTest_Result)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@Remark", fabricCrkShrkTest_Result.Main.FirLaboratoryRemark);
            listPar.Add("@POID", fabricCrkShrkTest_Result.Main.POID);

            string sqlUpdatePO = @"
update PO set FirLaboratoryRemark = @Remark
where ID = @POID
";

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory set  ReceiveSampleDate = @ReceiveSampleDate,
                            nonCrocking = @nonCrocking,
                            nonHeat = @nonHeat,
                            nonWash = @nonWash
where   ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdatePO, listPar);

                foreach (FabricCrkShrkTest_Detail fabricCrkShrkTest_Detail in fabricCrkShrkTest_Result.Details)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();
                    listDetailPar.Add("@ID", fabricCrkShrkTest_Detail.ID);
                    listDetailPar.Add("@ReceiveSampleDate", fabricCrkShrkTest_Detail.ReceiveSampleDate);
                    listDetailPar.Add("@nonCrocking", fabricCrkShrkTest_Detail.NonCrocking);
                    listDetailPar.Add("@nonHeat", fabricCrkShrkTest_Detail.NonHeat);
                    listDetailPar.Add("@nonWash", fabricCrkShrkTest_Detail.NonWash);

                    ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listDetailPar);
                }

                transaction.Complete();
            }

        }


        public FabricCrkShrkTestCrocking_Main GetFabricCrockingTest_Main(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrockingTest_Main = @"

select	[POID] = f.POID,
        [SEQ] = Concat(f.Seq1, ' ', f.Seq2),
        [ColorID] = psd.ColorID,
        [ArriveQty] = f.ArriveQty,
        [WhseArrival] = r.WhseArrival,
        [ExportID] = r.ExportID,
        [Supp] = Concat(f.SuppID, s.AbbEn),
        [Crocking] = fl.Crocking,
        [CrockingDate] = fl.CrockingDate,
        [StyleID] = o.StyleID,
        [SCIRefno] = f.SCIRefno,
        [Name] = (select Name from pass1 where ID = fl.CrockingInspector),
        [BrandID] = o.BrandID,
        [Refno] = f.Refno,
        [NonCrocking] = fl.NonCrocking,
        [DescDetail] = fab.DescDetail,
        [CrockingRemark] = fl.CrockingRemark,
        [CrockingEncdoe] = fl.CrockingEncode,
        [CrockingTestBeforePicture] = fl.CrockingTestBeforePicture,
        [CrockingTestAfterPicture] = fl.CrockingTestAfterPicture
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join Supp s with (nolock) on s.ID = f.SuppID
left join Orders o with (nolock) on o.ID = f.POID
left join Fabric fab with (nolock) on fab.SCIRefno = f.SCIRefno
where f.ID = @ID
";

            IList<FabricCrkShrkTestCrocking_Main> listResult = ExecuteList<FabricCrkShrkTestCrocking_Main>(CommandType.Text, sqlGetFabricCrockingTest_Main, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult[0];
        }

        public List<FabricCrkShrkTestCrocking_Detail> GetFabricCrockingTest_Detail(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrockingTest_Detail = @"

select	[Roll] = flc.Roll,
        [Dyelot] = flc.Dyelot,
        [Result] = flc.Result,
        [DryScale] = flc.DryScale,
        [ResultDry] = flc.ResultDry,
        [DryScale_Weft] = flc.DryScale_Weft,
        [ResultDry_Weft] = flc.ResultDry_Weft,
        [WetScale] = flc.WetScale,
        [ResultWet] = flc.ResultWet,
        [WetScale_Weft] = flc.WetScale_Weft,
        [ResultWet_Weft] = flc.ResultWet_Weft,
        [Inspdate] = flc.Inspdate,
        [Inspector] = flc.Inspector,
        [Name] = (select Name from pass1 where ID = flc.Inspector),
        [Remark] = flc.Remark,
        [LastUpdate] = Concat(LastUpdateName.val, ' - ', isnull(Format(flc.EditDate, 'yyyy/MM/dd HH:mm:ss'), Format(flc.AddDate, 'yyyy/MM/dd HH:mm:ss')))
from FIR_Laboratory_Crocking flc with (nolock)
outer apply (select [val] = Name_Extno from View_ShowName where ID = iif(isnull(flc.EditName, '') = '', flc.AddName, flc.EditName)) LastUpdateName
where flc.ID = @ID
";

            return ExecuteList<FabricCrkShrkTestCrocking_Detail>(CommandType.Text, sqlGetFabricCrockingTest_Detail, listPar).ToList();

        }

        public int GetCrockingTestOption(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);
            int crockingTestOption = 0;

            string sqlGetCrockingTestOption = @"
select  CrockingTestOption
from    QABrandSetting with (nolock)
where   BrandID = (select BrandID from orders with (nolock) where ID = (select POID from FIR with (nolock) where ID = @ID))
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetCrockingTestOption, listPar);

            if (dtResult.Rows.Count > 0)
            {
                crockingTestOption = int.Parse(dtResult.Rows[0]["CrockingTestOption"].ToString());
            }

            return crockingTestOption;
        }

        public void UpdateFabricCrockingTestDetail(FabricCrkShrkTestCrocking_Result fabricCrkShrkTestCrocking_Result, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", fabricCrkShrkTestCrocking_Result.ID);
            listPar.Add("@CrockingRemark", fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingRemark);

            string sqlUpdateCrocking = @"
update  FIR_Laboratory set CrockingRemark = @CrockingRemark
where   ID = @ID 
";


            List<FabricCrkShrkTestCrocking_Detail> oldCrockingData = GetFabricCrockingTest_Detail(fabricCrkShrkTestCrocking_Result.ID);

            List<FabricCrkShrkTestCrocking_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<FabricCrkShrkTestCrocking_Detail>(
                    fabricCrkShrkTestCrocking_Result.Crocking_Detail,
                    oldCrockingData,
                    "Roll,Dyelot",
                    "Result,DryScale,ResultDry,DryScale_Weft,ResultDry_Weft,WetScale,ResultWet,WetScale_Weft,ResultWet_Weft,Inspdate,Inspector,Remark");



            string sqlInsertDetail = @"
insert into FIR_Laboratory_Crocking(
ID              ,
Roll            ,
Dyelot          ,
DryScale        ,
WetScale        ,
Inspdate        ,
Inspector       ,
Result          ,
Remark          ,
AddName         ,
AddDate         ,
ResultDry       ,
ResultWet       ,
DryScale_Weft   ,
WetScale_Weft   ,
ResultDry_Weft  ,
ResultWet_Weft
)
values
(
@ID              ,
@Roll            ,
@Dyelot          ,
@DryScale        ,
@WetScale        ,
@Inspdate        ,
@Inspector       ,
@Result          ,
@Remark          ,
@AddName         ,
getDate()         ,
@ResultDry       ,
@ResultWet       ,
@DryScale_Weft   ,
@WetScale_Weft   ,
@ResultDry_Weft  ,
@ResultWet_Weft
)

";

            string sqlDeleteDetail = @"
delete  FIR_Laboratory_Crocking
where   ID = @ID and
        Roll = @Roll and
        Dyelot = @Dyelot
";

            string sqlUpdateDetail = @"
update  FIR_Laboratory_Crocking set DryScale       = @DryScale      ,
                                    WetScale       = @WetScale      ,
                                    Inspdate       = @Inspdate      ,
                                    Inspector      = @Inspector     ,
                                    Result         = @Result        ,
                                    Remark         = @Remark        ,
                                    EditName       = @EditName      ,
                                    EditDate       = getDate()      ,
                                    ResultDry      = @ResultDry     ,
                                    ResultWet      = @ResultWet     ,
                                    DryScale_Weft  = @DryScale_Weft ,
                                    WetScale_Weft  = @WetScale_Weft ,
                                    ResultDry_Weft = @ResultDry_Weft,
                                    ResultWet_Weft = @ResultWet_Weft

        where   ID = @ID and
                Roll = @Roll and
                Dyelot = @Dyelot
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdateCrocking, listPar);
                foreach (FabricCrkShrkTestCrocking_Detail detailItem in needUpdateDetailList)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();

                    switch (detailItem.StateType)
                    {
                        case DatabaseObject.Public.CompareStateType.Add:
                            listDetailPar.Add("@ID", fabricCrkShrkTestCrocking_Result.ID);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@DryScale", detailItem.DryScale);
                            listDetailPar.Add("@WetScale", detailItem.WetScale);
                            listDetailPar.Add("@Inspdate", detailItem.Inspdate);
                            listDetailPar.Add("@Inspector", detailItem.Inspector);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@Remark", detailItem.Remark);
                            listDetailPar.Add("@AddName", userID);
                            listDetailPar.Add("@ResultDry", detailItem.ResultDry);
                            listDetailPar.Add("@ResultWet", detailItem.ResultWet);
                            listDetailPar.Add("@DryScale_Weft", detailItem.DryScale_Weft);
                            listDetailPar.Add("@WetScale_Weft", detailItem.WetScale_Weft);
                            listDetailPar.Add("@ResultDry_Weft", detailItem.ResultDry_Weft);
                            listDetailPar.Add("@ResultWet_Weft", detailItem.ResultWet_Weft);

                            ExecuteNonQuery(CommandType.Text, sqlInsertDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Edit:
                            listDetailPar.Add("@ID", fabricCrkShrkTestCrocking_Result.ID);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@DryScale", detailItem.DryScale);
                            listDetailPar.Add("@WetScale", detailItem.WetScale);
                            listDetailPar.Add("@Inspdate", detailItem.Inspdate);
                            listDetailPar.Add("@Inspector", detailItem.Inspector);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@Remark", detailItem.Remark);
                            listDetailPar.Add("@EditName", userID);
                            listDetailPar.Add("@ResultDry", detailItem.ResultDry);
                            listDetailPar.Add("@ResultWet", detailItem.ResultWet);
                            listDetailPar.Add("@DryScale_Weft", detailItem.DryScale_Weft);
                            listDetailPar.Add("@WetScale_Weft", detailItem.WetScale_Weft);
                            listDetailPar.Add("@ResultDry_Weft", detailItem.ResultDry_Weft);
                            listDetailPar.Add("@ResultWet_Weft", detailItem.ResultWet_Weft);

                            ExecuteNonQuery(CommandType.Text, sqlUpdateDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Delete:
                            listDetailPar.Add("@ID", fabricCrkShrkTestCrocking_Result.ID);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);

                            ExecuteNonQuery(CommandType.Text, sqlDeleteDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.None:
                            break;
                        default:
                            break;
                    }
                }

                transaction.Complete();
            }
        }

        public string GetTestPOID()
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlGet = @"
select top 1 POID
from    FIR_Laboratory
where   Result <> '' and CrockingDate between getdate() - 160 and getdate() + 160

";
            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGet, listPar).Rows[0][0].ToString();
        }

        public long GetTestFIRID()
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlGet = @"
select top 1 ID
from    FIR_Laboratory
where   Result <> '' and CrockingDate between getdate() - 160 and getdate() + 160

";
            return (long)ExecuteDataTableByServiceConn(CommandType.Text, sqlGet, listPar).Rows[0][0];
        }

        public void EncodeFabricCrocking(long ID, string testResult, DateTime? crockingDate, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);
            listPar.Add("@testResult", testResult);
            listPar.Add("@CrockingDate", crockingDate);
            listPar.Add("@userID", userID);

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory  set Crocking = @testResult,
                            CrockingDate = @CrockingDate,
                            CrockingEncode = 1,
                            CrockingInspector = @userID
     where  ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listPar);
                transaction.Complete();
            }
        }

        public DataTable GetCrockingFailMailContentData(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetData = @"
select	[SP#] = f.POID,
        [Style] = o.StyleID,
        [Brand] = o.BrandID,
        [Season] = o.SeasonID,
        [SEQ] = Concat(f.Seq1, ' ', f.Seq2),
        [WK#] = r.ExportID,
        [Arrive WH Date] = Format(r.WhseArrival, 'yyyy/MM/dd'),
        [SCI Refno] = f.SCIRefno,
        [Refno] = f.Refno,
        [Color] = psd.ColorID,
        [Supplier] = Concat(f.SuppID, s.AbbEn),
        [Arrive Qty] = f.ArriveQty,
        [Crocking Result] = fl.Crocking,
        [Crocking Last Test Date] = Format(fl.CrockingDate, 'yyyy/MM/dd'),
        [Crocking Remark] = fl.CrockingRemark
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join Supp s with (nolock) on s.ID = f.SuppID
left join Orders o with (nolock) on o.ID = f.POID
left join Fabric fab with (nolock) on fab.SCIRefno = f.SCIRefno
where f.ID = @ID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public void AmendFabricCrocking(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory  set Crocking = '',
                            CrockingDate = null,
                            CrockingEncode = 0,
                            CrockingInspector = ''
     where  ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listPar);
                transaction.Complete();
            }
        }

        public DataTable GetCrockingDetailForReport(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetData = @"
select	[Roll] = flc.Roll,
        [Dyelot] = flc.Dyelot,
        [DryScale] = flc.DryScale,
        [WetScale] = flc.WetScale,
        [Result] = flc.Result,
        [Inspdate] = flc.Inspdate,
        [Inspector] = flc.Inspector,
        [Remark] = flc.Remark,
        [LastUpdate] = Concat(LastUpdateName.val, ' - ', isnull(Format(flc.EditDate, 'yyyy/MM/dd HH:mm:ss'), Format(flc.AddDate, 'yyyy/MM/dd HH:mm:ss'))),
        [DryScale_Weft] = flc.DryScale_Weft,
        [WetScale_Weft] = flc.WetScale_Weft,
        [ResultDry] = flc.ResultDry
from FIR_Laboratory_Crocking flc with (nolock)
outer apply (select [val] = Name_Extno from View_ShowName where ID = iif(isnull(flc.EditName, '') = '', flc.AddName, flc.EditName)) LastUpdateName
where flc.ID = @ID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        #endregion

    }
}
