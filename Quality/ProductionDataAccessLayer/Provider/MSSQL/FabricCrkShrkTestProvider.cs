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
        [TestBeforePicture] = (select top 1 TestBeforePicture from SciPMSFile_Oven oi with (nolock) where o.ID = oi.ID) ,
        [TestAfterPicture] =  (select top 1 TestAfterPicture from SciPMSFile_Oven oi with (nolock) where o.ID = oi.ID) 
from Oven o with (nolock)
left join pass1 pass1AddName WITH(NOLOCK) on o.AddName = pass1AddName.ID
left join pass1 pass1EditName WITH(NOLOCK) on o.EditName = pass1EditName.ID
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
        [ColorID] = pc.SpecValue,
        [Result] = od.Result,
        [ChangeScale] = od.changeScale,
        [ResultChange] = od.ResultChange,
        [StainingScale] = od.StainingScale,
        [ResultStain] = od.ResultStain,
        [Remark] = od.Remark,
        [LastUpdate] = Concat(od.EditName, '-', pass1EditName.Name, ' ', pass1EditName.Extno),
        [Temperature] = od.Temperature,
        [Time] = od.Time
from Oven_Detail od with (nolock)
inner join Oven o with (nolock) on o.ID = od.ID
left join PO_Supp_Detail psd with (nolock) on o.POID = psd.ID and od.SEQ1 = psd.SEQ1 and od.SEQ2 = psd.SEQ2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
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
		[CompletionDate] = iif(p.LabOvenPercent >= 100, FIR_CompletionDate.val, null),
		[FIRLabInspPercent] = p.FIRLabInspPercent,
        [complete] = p.complete,
		[FirLaboratoryRemark] = p.FirLaboratoryRemark,
		[CreateBy] = Concat(p.AddName, '-', pass1AddName.Name, ' ', Format(p.AddDate, 'yyyy/MM/dd HH:mm:ss')),
		[EditBy] = Concat(p.EditName, '-', pass1EditName.Name, ' ', Format(p.EditDate, 'yyyy/MM/dd HH:mm:ss'))
from PO p with (nolock)
inner join Orders o with (nolock) on p.ID = o.ID
left join pass1 pass1AddName WITH(NOLOCK) on p.AddName = pass1AddName.ID
left join pass1 pass1EditName WITH(NOLOCK) on p.EditName = pass1EditName.ID
outer apply (    select [val] = max(CompletionDate)
                from    (
                 select [CompletionDate] = Max(CrockingDate) from FIR_Laboratory WITH(NOLOCK) where POID = p.ID
                 union all
                 select [CompletionDate] = Max(HeatDate) from FIR_Laboratory WITH(NOLOCK) where POID = p.ID
                 union all
                 select [CompletionDate] = Max(WashDate) from FIR_Laboratory WITH(NOLOCK) where POID = p.ID
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
		fl.ReportNo,
        [Seq] = Concat(f.Seq1, ' ', f.Seq2),
        [WKNo] = r.ExportID,
        [WhseArrival] = r.WhseArrival,
        [SCIRefno] = f.SCIRefno,
        [Refno] = f.Refno,
        [ColorID] = pc.SpecValue,
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
        [NonIron] = fl.nonIron,
        [Iron] = fl.Iron,
        [IronDate] = fl.IronDate,
        [IronRemark] = fl.IronRemark,
        [NonWash] = fl.nonWash,
        [Wash] = fl.Wash,
        [WashDate] = fl.WashDate,
        [WashRemark] = fl.WashRemark,
        [ReceivingID] = f.ReceivingID
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Supp s with (nolock) on s.ID = f.SuppID
where f.POID = @POID
";

            fabricCrkShrkTest_Result.Details = ExecuteList<FabricCrkShrkTest_Detail>(CommandType.Text, sqlGetDetails, listPar).ToList();

            return fabricCrkShrkTest_Result;
        }


        public void SaveFabricCrkShrkTest_Main(FabricCrkShrkTest_Result fabricCrkShrkTest_Result)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@Remark", fabricCrkShrkTest_Result.Main.FirLaboratoryRemark ?? string.Empty);
            listPar.Add("@POID", fabricCrkShrkTest_Result.Main.POID);

            string sqlUpdatePO = @"
update PO set FirLaboratoryRemark = @Remark
where ID = @POID

exec UpdateInspPercent 'FIRLab',@POID
";

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory set  ReceiveSampleDate = @ReceiveSampleDate,
                            nonCrocking = @nonCrocking,
                            nonHeat = @nonHeat,
                            nonIron = @nonIron,
                            nonWash = @nonWash
where   ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}
";

            using (TransactionScope transaction = new TransactionScope())
            {

                if (fabricCrkShrkTest_Result.Details != null)
                {
                    foreach (FabricCrkShrkTest_Detail fabricCrkShrkTest_Detail in fabricCrkShrkTest_Result.Details)
                    {
                        SQLParameterCollection listDetailPar = new SQLParameterCollection();
                        listDetailPar.Add("@ID", fabricCrkShrkTest_Detail.ID);
                        listDetailPar.Add("@ReceiveSampleDate", fabricCrkShrkTest_Detail.ReceiveSampleDate);
                        listDetailPar.Add("@nonCrocking", fabricCrkShrkTest_Detail.NonCrocking);
                        listDetailPar.Add("@nonHeat", fabricCrkShrkTest_Detail.NonHeat);
                        listDetailPar.Add("@NonIron", fabricCrkShrkTest_Detail.NonIron);
                        listDetailPar.Add("@nonWash", fabricCrkShrkTest_Detail.NonWash);

                        ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listDetailPar);
                    }
                }

                ExecuteNonQuery(CommandType.Text, sqlUpdatePO, listPar);

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
        [ColorID] = pc.SpecValue,
        [ArriveQty] = f.ArriveQty,
        [WhseArrival] = r.WhseArrival,
        [ExportID] = r.ExportID,
        [Supp] = Concat(f.SuppID, s.AbbEn),
        [Crocking] = fl.Crocking,
        [CrockingDate] = fl.CrockingDate,
        [StyleID] = o.StyleID,
        [SCIRefno] = f.SCIRefno,
        [Name] = (select Name from pass1 WITH(NOLOCK) where ID = fl.CrockingInspector),
        [BrandID] = o.BrandID,
        [Refno] = f.Refno,
        [NonCrocking] = fl.NonCrocking,
        [DescDetail] = fab.DescDetail,
        [CrockingRemark] = fl.CrockingRemark,
        [CrockingEncdoe] = fl.CrockingEncode,
        [CrockingTestBeforePicture] = (select top 1 CrockingTestBeforePicture from SciPMSFile_FIR_Laboratory fli WITH(NOLOCK) where fli.ID = fl.ID ),
        [CrockingTestAfterPicture] = (select top 1 CrockingTestAfterPicture from SciPMSFile_FIR_Laboratory fli WITH(NOLOCK) where fli.ID = fl.ID )
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
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
        [Name] = (select Concat(Name, ' Ext.', ExtNo) from pass1 where ID = flc.Inspector),
        [Remark] = flc.Remark,
        [LastUpdate] = Concat(LastUpdateName.val, ' - ', isnull(Format(flc.EditDate, 'yyyy/MM/dd HH:mm:ss'), Format(flc.AddDate, 'yyyy/MM/dd HH:mm:ss')))
from FIR_Laboratory_Crocking flc with (nolock)
outer apply (select [val] = Name_Extno from View_ShowName where ID = iif(isnull(flc.EditName, '') = '', flc.AddName, flc.EditName)) LastUpdateName
where flc.ID = @ID
";

            return ExecuteList<FabricCrkShrkTestCrocking_Detail>(CommandType.Text, sqlGetFabricCrockingTest_Detail, listPar).ToList();

        }

        public List<Crocking_Excel> CrockingTest_ToExcel(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sql = @"
select	SubmitDate = fl.CrockingDate
		,o.SeasonID
		,o.BrandID
		,o.StyleID
		,sa.Article
		,f.POID
		,fd.Roll
		,fd.Dyelot
		,SCIRefno_Color = f.SCIRefno + ' ' + pc.SpecValue
		,Color = pc.SpecValue
		,fd.DryScale
		,fd.DryScale_Weft
		,fd.WetScale
		,fd.WetScale_Weft
		,fd.ResultDry
		,fd.ResultDry_Weft
		,fd.ResultWet
		,fd.ResultWet_Weft
		,fd.Remark
		,fd.Inspector
		,CrockingTestBeforePicture = (select top 1 CrockingTestBeforePicture from SciPMSFile_FIR_Laboratory fli WITH(NOLOCK) where fli.ID = fl.ID)
        ,CrockingTestAfterPicture = (select top 1 CrockingTestAfterPicture from SciPMSFile_FIR_Laboratory fli WITH(NOLOCK) where fli.ID = fl.ID)
		,f.ID
        ,fl.ReportNo
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
inner join FIR_Laboratory_Crocking fd WITH(NOLOCK) on fd.id = fl.id
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Orders o with (nolock) on o.ID = f.POID
left join Style_Article sa ON o.StyleUkey=sa.StyleUkey
where f.ID = @ID
";

            return ExecuteList<Crocking_Excel>(CommandType.Text, sql, listPar).ToList();

        }

        public List<Crocking_Excel> CrockingTest_ToExcel_Head(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sql = @"
select fl.ReportNo
	,f.POID
	,Article=Article.Val
	,SubmitDate = fl.CrockingDate
	,o.SeasonID
	,o.BrandID
	,o.StyleID
	,SCIRefno_Color = f.SCIRefno + ' ' + pc.SpecValue
	,Color = pc.SpecValue
	,Inspector = LabTech.Val
	,CrockingTestBeforePicture = (select top 1 CrockingTestBeforePicture from SciPMSFile_FIR_Laboratory fli WITH(NOLOCK) where fli.ID = fl.ID)
    ,CrockingTestAfterPicture = (select top 1 CrockingTestAfterPicture from SciPMSFile_FIR_Laboratory fli WITH(NOLOCK) where fli.ID = fl.ID)
from FIR_Laboratory fl
inner join Orders o with (nolock) on o.ID = fl.POID
inner join FIR f with (nolock) on f.ID = fl.ID
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
outer apply(
	select Val=STUFF((
		select DISTINCT ',' + sa.Article
		from Style_Article sa WITH(NOLOCK)
		where o.StyleUkey=sa.StyleUkey
		FOR XML PATH('')
	),1,1,'')
)Article
outer apply(
	select Val=STUFF((
		select DISTINCT ',' + fd.Inspector
		from FIR_Laboratory_Crocking fd WITH(NOLOCK) 
		where fd.id = fl.id
		FOR XML PATH('')
	),1,1,'')
)LabTech
where fl.ID = @ID
";

            return ExecuteList<Crocking_Excel>(CommandType.Text, sql, listPar).ToList();

        }
        public List<Crocking_Excel> CrockingTest_ToExcel_Body(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sql = @"
select   fd.Roll
		,fd.Dyelot
		,fd.DryScale
		,fd.DryScale_Weft
		,fd.WetScale
		,fd.WetScale_Weft
		,fd.ResultDry
		,fd.ResultDry_Weft
		,fd.ResultWet
		,fd.ResultWet_Weft
		,fd.Remark
from FIR_Laboratory fl
inner join FIR_Laboratory_Crocking fd WITH(NOLOCK) on fd.id = fl.id
where fl.ID = @ID
";

            return ExecuteList<Crocking_Excel>(CommandType.Text, sql, listPar).ToList();

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
            listPar.Add("@CrockingTestBeforePicture", fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestBeforePicture);
            listPar.Add("@CrockingTestAfterPicture", fabricCrkShrkTestCrocking_Result.Crocking_Main.CrockingTestAfterPicture);

            string sqlUpdateCrocking = @"
SET XACT_ABORT ON
-- 2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
update  FIR_Laboratory set  CrockingRemark = @CrockingRemark
where   ID = @ID 

if exists( select 1 from SciPMSFile_FIR_Laboratory where   ID = @ID )
begin
    update  SciPMSFile_FIR_Laboratory set  CrockingTestBeforePicture = @CrockingTestBeforePicture,
                                CrockingTestAfterPicture = @CrockingTestAfterPicture
    where   ID = @ID 
END
else
begin
    insert into SciPMSFile_FIR_Laboratory (ID  ,CrockingTestBeforePicture  ,CrockingTestAfterPicture)
    VALUES(@ID  ,@CrockingTestBeforePicture  ,@CrockingTestAfterPicture )
end
";


            List<FabricCrkShrkTestCrocking_Detail> oldCrockingData = GetFabricCrockingTest_Detail(fabricCrkShrkTestCrocking_Result.ID);

            List<FabricCrkShrkTestCrocking_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<FabricCrkShrkTestCrocking_Detail>(
                    fabricCrkShrkTestCrocking_Result.Crocking_Detail,
                    oldCrockingData,
                    "Roll,Dyelot",
                    "Result,DryScale,ResultDry,DryScale_Weft,ResultDry_Weft,WetScale,ResultWet,WetScale_Weft,ResultWet_Weft,Inspdate,Inspector,Remark");


            string NewReportNo = GetID(fabricCrkShrkTestCrocking_Result.MDivisionID + "FT", "FIR_Laboratory", DateTime.Today, 2, "ReportNo");

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
;
UPDATE FIR_Laboratory
SET ReportNo = @ReportNo
WHERE ReportNo = '' AND ID= @ID
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
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@AddName", userID);
                            listDetailPar.Add("@ResultDry", detailItem.ResultDry);
                            listDetailPar.Add("@ResultWet", detailItem.ResultWet);
                            listDetailPar.Add("@DryScale_Weft", detailItem.DryScale_Weft);
                            listDetailPar.Add("@WetScale_Weft", detailItem.WetScale_Weft);
                            listDetailPar.Add("@ResultDry_Weft", detailItem.ResultDry_Weft);
                            listDetailPar.Add("@ResultWet_Weft", detailItem.ResultWet_Weft);
                            listDetailPar.Add("@ReportNo", NewReportNo);

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
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
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

                string UpdateInspPercent = $@"
DECLARE @POID as varchar(15)= (SELECT TOP 1  POID from FIR_Laboratory where ID = @ID)
exec UpdateInspPercent 'FIRLab', @POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }

        public string GetTestPOID(string where)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlGet = $@"
select top 1 POID
from    FIR_Laboratory WITH(NOLOCK)
where   Result <> '' and CrockingDate between getdate() - 160 and getdate() + 160
{where}
";
            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGet, listPar).Rows[0][0].ToString();
        }

        public long GetTestFIRID(string where)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string sqlGet = $@"
select top 1 ID
from    FIR_Laboratory WITH(NOLOCK)
where   Result <> '' and CrockingDate between getdate() - 160 and getdate() + 160
{where}
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

declare @POID varchar(13)

select @POID = POID from FIR_Laboratory WITH(NOLOCK) where ID = @ID

exec UpdateInspPercent 'FIRLab',@POID 
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
        [Color] = pc.SpecValue,
        [Supplier] = Concat(f.SuppID, s.AbbEn),
        [Arrive Qty] = f.ArriveQty,
        [Crocking Result] = fl.Crocking,
        [Crocking Last Test Date] = Format(fl.CrockingDate, 'yyyy/MM/dd'),
        [Crocking Remark] = fl.CrockingRemark
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
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

declare @POID varchar(13)

select @POID = POID from FIR_Laboratory WITH(NOLOCK) where ID = @ID

exec UpdateInspPercent 'FIRLab',@POID 
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

        public DataTable GetCrockingArticleForPdfReport(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetData = @"
SELECT distinct fd.InspDate,
                oc.Article,
                fd.DryScale,
                fd.ResultDry, 
                fd.DryScale_Weft, 
                fd.ResultDry_Weft,
                fd.WetScale_Weft,
                fd.ResultWet_Weft,
                fd.WetScale,
                fd.ResultWet,
                fd.Remark,
                fd.Inspector,
                fd.Roll,
                fd.Dyelot,
                fd.Result,
                a.Name
FROM Order_BOF bof WITH(NOLOCK)
inner join PO_Supp_Detail p WITH(NOLOCK) on p.id=bof.id and bof.SCIRefno=p.SCIRefno
inner join Order_ColorCombo OC WITH(NOLOCK) on oc.id=p.id and oc.FabricCode=bof.FabricCode
inner join orders o WITH(NOLOCK) on o.id = bof.id
inner join FIR_Laboratory f WITH(NOLOCK) on f.poid = o.poid and f.seq1 = p.seq1 and f.seq2 = p.seq2
inner join FIR_Laboratory_Crocking fd WITH(NOLOCK) on fd.id = f.id
outer apply
(
	select Name = stuff((
		select concat(',',Name)
		from pass1  WITH(NOLOCK)
		where id = fd.Inspector
		for xml path('')
	),1,1,'')
)a
where  f.ID = @ID
order by fd.InspDate,oc.article
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public FabricCrkShrkTestHeat_Main GetFabricHeatTest_Main(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrkShrkTestHeat_Main = @"

select	[POID] = f.POID,
        [SEQ] = Concat(f.Seq1, ' ', f.Seq2),
        [ColorID] = pc.SpecValue,
        [ArriveQty] = f.ArriveQty,
        [WhseArrival] = r.WhseArrival,
        [ExportID] = r.ExportID,
        [Supp] = Concat(f.SuppID, s.AbbEn),
        [Heat] = fl.Heat,
        [HeatDate] = fl.HeatDate,
        [StyleID] = o.StyleID,
        [SCIRefno] = f.SCIRefno,
        [Name] = (select Name from pass1 WITH(NOLOCK) where ID = fl.HeatInspector),
        [BrandID] = o.BrandID,
        [Refno] = f.Refno,
        [NonHeat] = fl.NonHeat,
        [DescDetail] = fab.DescDetail,
        [HeatRemark] = fl.HeatRemark,
        [HeatEncode] = fl.HeatEncode,
        [HeatTestBeforePicture] =  (select top 1 HeatTestBeforePicture from SciPMSFile_FIR_Laboratory fli WITH (NOLOCK) where fli.ID = fl.ID ),  
        [HeatTestAfterPicture] = (select top 1 HeatTestAfterPicture from SciPMSFile_FIR_Laboratory fli WITH (NOLOCK) where fli.ID = fl.ID),
        [ReportNo] = fl.ReportNo
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Supp s with (nolock) on s.ID = f.SuppID
left join Orders o with (nolock) on o.ID = f.POID
left join Fabric fab with (nolock) on fab.SCIRefno = f.SCIRefno
where f.ID = @ID
";

            IList<FabricCrkShrkTestHeat_Main> listResult = ExecuteList<FabricCrkShrkTestHeat_Main>(CommandType.Text, sqlGetFabricCrkShrkTestHeat_Main, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult[0];
        }

        public List<FabricCrkShrkTestHeat_Detail> GetFabricHeatTest_Detail(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrkShrkTestHeat_Detail = @"

select	[Roll] = flc.Roll,
        [Dyelot] = flc.Dyelot,
        [HorizontalOriginal] = flc.HorizontalOriginal,
        [VerticalOriginal] = flc.VerticalOriginal,
        [Result] = flc.Result,
        [HorizontalTest1] = flc.HorizontalTest1,
        [HorizontalTest2] = flc.HorizontalTest2,
        [HorizontalTest3] = flc.HorizontalTest3,
        [HorizontalRate] = flc.HorizontalRate,
        [HorizontalAverage] = Cast(Round((isnull(flc.HorizontalTest1, 0) + isnull(flc.HorizontalTest2, 0)  + isnull(flc.HorizontalTest3, 0)) / 3.0, 2) as numeric(5, 2)),
        [VerticalTest1] = flc.VerticalTest1,
        [VerticalTest2] = flc.VerticalTest2,
        [VerticalTest3] = flc.VerticalTest3,
        [VerticalRate] = flc.VerticalRate,
        [VerticalAverage] = Cast(Round((isnull(flc.VerticalTest1, 0) + isnull(flc.VerticalTest2, 0)  + isnull(flc.VerticalTest3, 0)) / 3.0, 2) as numeric(5, 2)),
        [Inspdate] = flc.Inspdate,
        [Inspector] = flc.Inspector,
        [Name] = (select Concat(Name, 'Ext.', ExtNo) from pass1 WITH(NOLOCK) where ID = flc.Inspector),
        [Remark] = flc.Remark,
        [LastUpdate] = Concat(LastUpdateName.val, ' - ', isnull(Format(flc.EditDate, 'yyyy/MM/dd HH:mm:ss'), Format(flc.AddDate, 'yyyy/MM/dd HH:mm:ss')))
from FIR_Laboratory_Heat flc with (nolock)
outer apply (select [val] = Name_Extno from View_ShowName where ID = iif(isnull(flc.EditName, '') = '', flc.AddName, flc.EditName)) LastUpdateName
where flc.ID = @ID
";

            return ExecuteList<FabricCrkShrkTestHeat_Detail>(CommandType.Text, sqlGetFabricCrkShrkTestHeat_Detail, listPar).ToList();
        }

        public void UpdateFabricHeatTestDetail(FabricCrkShrkTestHeat_Result fabricCrkShrkTestHeat_Result, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", fabricCrkShrkTestHeat_Result.ID);
            listPar.Add("@HeatRemark", fabricCrkShrkTestHeat_Result.Heat_Main.HeatRemark);
            listPar.Add("@HeatTestBeforePicture", fabricCrkShrkTestHeat_Result.Heat_Main.HeatTestBeforePicture);
            listPar.Add("@HeatTestAfterPicture", fabricCrkShrkTestHeat_Result.Heat_Main.HeatTestAfterPicture);

            string sqlUpdateCrocking = @"
SET XACT_ABORT ON
-----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
update  FIR_Laboratory set  HeatRemark = @HeatRemark
where   ID = @ID 
;
if exists(
    select 1 from SciPMSFile_FIR_Laboratory where  ID = @ID 
)
begin
    update  SciPMSFile_FIR_Laboratory set HeatTestBeforePicture = @HeatTestBeforePicture,
                                HeatTestAfterPicture = @HeatTestAfterPicture
    where   ID = @ID 
end
else
begin
    insert into SciPMSFile_FIR_Laboratory (ID,HeatTestBeforePicture,HeatTestAfterPicture)
    values (@ID ,@HeatTestBeforePicture ,@HeatTestAfterPicture )
end
";


            List<FabricCrkShrkTestHeat_Detail> oldHeatData = GetFabricHeatTest_Detail(fabricCrkShrkTestHeat_Result.ID);

            List<FabricCrkShrkTestHeat_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<FabricCrkShrkTestHeat_Detail>(
                    fabricCrkShrkTestHeat_Result.Heat_Detail,
                    oldHeatData,
                    "Roll,Dyelot",
                    "HorizontalOriginal,VerticalOriginal,Result,HorizontalTest1,HorizontalTest2,HorizontalTest3,VerticalTest1,VerticalTest2,VerticalTest3,Inspdate,Inspector,Remark");

            string NewReportNo = GetID(fabricCrkShrkTestHeat_Result.MDivisionID + "FT", "FIR_Laboratory", DateTime.Today, 2, "ReportNo");

            string sqlInsertDetail = @"
insert into FIR_Laboratory_Heat(
ID                   ,
Roll                 ,
Dyelot               ,
Inspdate             ,
Inspector            ,
Result               ,
Remark               ,
AddName              ,
AddDate              ,
HorizontalRate       ,
HorizontalOriginal   ,
HorizontalTest1      ,
HorizontalTest2      ,
HorizontalTest3      ,
VerticalRate         ,
VerticalOriginal     ,
VerticalTest1        ,
VerticalTest2        ,
VerticalTest3

)
values
(
@ID                   ,
@Roll                 ,
@Dyelot               ,
@Inspdate             ,
@Inspector            ,
@Result               ,
@Remark               ,
@AddName              ,
getDate()              ,
@HorizontalRate       ,
@HorizontalOriginal   ,
@HorizontalTest1      ,
@HorizontalTest2      ,
@HorizontalTest3      ,
@VerticalRate         ,
@VerticalOriginal     ,
@VerticalTest1        ,
@VerticalTest2        ,
@VerticalTest3
)
;
UPDATE FIR_Laboratory
SET ReportNo = @ReportNo
WHERE ReportNo = '' AND ID= @ID
";

            string sqlDeleteDetail = @"
delete  FIR_Laboratory_Heat
where   ID = @ID and
        Roll = @Roll and
        Dyelot = @Dyelot
";

            string sqlUpdateDetail = @"
update  FIR_Laboratory_Heat set Inspdate            = @Inspdate             ,
                                Inspector           = @Inspector            ,
                                Result              = @Result               ,
                                Remark              = @Remark               ,
                                EditName            = @EditName             ,
                                EditDate            = getDate()             ,
                                HorizontalRate      = @HorizontalRate       ,
                                HorizontalOriginal  = @HorizontalOriginal   ,
                                HorizontalTest1     = @HorizontalTest1      ,
                                HorizontalTest2     = @HorizontalTest2      ,
                                HorizontalTest3     = @HorizontalTest3      ,
                                VerticalRate        = @VerticalRate         ,
                                VerticalOriginal    = @VerticalOriginal     ,
                                VerticalTest1       = @VerticalTest1        ,
                                VerticalTest2       = @VerticalTest2        ,
                                VerticalTest3       = @VerticalTest3     
        where   ID = @ID and
                Roll = @Roll and
                Dyelot = @Dyelot
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdateCrocking, listPar);
                foreach (FabricCrkShrkTestHeat_Detail detailItem in needUpdateDetailList)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();

                    switch (detailItem.StateType)
                    {
                        case DatabaseObject.Public.CompareStateType.Add:
                            listDetailPar.Add("@ID", fabricCrkShrkTestHeat_Result.ID);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Inspdate", detailItem.Inspdate);
                            listDetailPar.Add("@Inspector", detailItem.Inspector);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@AddName", userID);
                            listDetailPar.Add("@HorizontalRate", detailItem.HorizontalRate);
                            listDetailPar.Add("@HorizontalOriginal", detailItem.HorizontalOriginal);
                            listDetailPar.Add("@HorizontalTest1", detailItem.HorizontalTest1);
                            listDetailPar.Add("@HorizontalTest2", detailItem.HorizontalTest2);
                            listDetailPar.Add("@HorizontalTest3", detailItem.HorizontalTest3);
                            listDetailPar.Add("@VerticalRate", detailItem.VerticalRate);
                            listDetailPar.Add("@VerticalOriginal", detailItem.VerticalOriginal);
                            listDetailPar.Add("@VerticalTest1", detailItem.VerticalTest1);
                            listDetailPar.Add("@VerticalTest2", detailItem.VerticalTest2);
                            listDetailPar.Add("@VerticalTest3", detailItem.VerticalTest3);
                            listDetailPar.Add("@ReportNo", NewReportNo);

                            ExecuteNonQuery(CommandType.Text, sqlInsertDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Edit:
                            listDetailPar.Add("@ID", fabricCrkShrkTestHeat_Result.ID);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Inspdate", detailItem.Inspdate);
                            listDetailPar.Add("@Inspector", detailItem.Inspector);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@EditName", userID);
                            listDetailPar.Add("@HorizontalRate", detailItem.HorizontalRate);
                            listDetailPar.Add("@HorizontalOriginal", detailItem.HorizontalOriginal);
                            listDetailPar.Add("@HorizontalTest1", detailItem.HorizontalTest1);
                            listDetailPar.Add("@HorizontalTest2", detailItem.HorizontalTest2);
                            listDetailPar.Add("@HorizontalTest3", detailItem.HorizontalTest3);
                            listDetailPar.Add("@VerticalRate", detailItem.VerticalRate);
                            listDetailPar.Add("@VerticalOriginal", detailItem.VerticalOriginal);
                            listDetailPar.Add("@VerticalTest1", detailItem.VerticalTest1);
                            listDetailPar.Add("@VerticalTest2", detailItem.VerticalTest2);
                            listDetailPar.Add("@VerticalTest3", detailItem.VerticalTest3);

                            ExecuteNonQuery(CommandType.Text, sqlUpdateDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Delete:
                            listDetailPar.Add("@ID", fabricCrkShrkTestHeat_Result.ID);
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

                string UpdateInspPercent = $@"
DECLARE @POID as varchar(15)= (SELECT TOP 1  POID from FIR_Laboratory where ID = @ID)
exec UpdateInspPercent 'FIRLab', @POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }

        public void EncodeFabricHeat(long ID, string testResult, DateTime? heatDate, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);
            listPar.Add("@testResult", testResult);
            listPar.Add("@HeatDate", heatDate);
            listPar.Add("@userID", userID);

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory  set Heat = @testResult,
                            HeatDate = @HeatDate,
                            HeatEncode  = 1,
                            HeatInspector = @userID
     where  ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}

declare @POID varchar(13)

select @POID = POID from FIR_Laboratory WITH(NOLOCK) where ID = @ID

exec UpdateInspPercent 'FIRLab',@POID 
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listPar);
                transaction.Complete();
            }
        }

        public DataTable GetHeatFailMailContentData(long ID)
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
        [Color] = pc.SpecValue,
        [Supplier] = Concat(f.SuppID, s.AbbEn),
        [Arrive Qty] = f.ArriveQty,
        [Heat Result] = fl.Heat,
        [Heat Last Test Date] = Format(fl.HeatDate, 'yyyy/MM/dd'),
        [Heat Remark] = fl.HeatRemark
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Supp s with (nolock) on s.ID = f.SuppID
left join Orders o with (nolock) on o.ID = f.POID
left join Fabric fab with (nolock) on fab.SCIRefno = f.SCIRefno
where f.ID = @ID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public void AmendFabricHeat(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory  set Heat = '',
                            HeatDate = null,
                            HeatEncode = 0,
                            HeatInspector = ''
     where  ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}

declare @POID varchar(13)

select @POID = POID from FIR_Laboratory WITH(NOLOCK) where ID = @ID

exec UpdateInspPercent 'FIRLab',@POID 
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listPar);
                transaction.Complete();
            }
        }

        public DataTable GetHeatDetailForReport(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrkShrkTestHeat_Detail = @"

select	[Roll] = flc.Roll,
        [Dyelot] = flc.Dyelot,
        [HorizontalOriginal] = flc.HorizontalOriginal,
        [VerticalOriginal] = flc.VerticalOriginal,
        [Result] = flc.Result,
        [HorizontalTest1] = flc.HorizontalTest1,
        [HorizontalTest2] = flc.HorizontalTest2,
        [HorizontalTest3] = flc.HorizontalTest3,
        [HorizontalRate] = flc.HorizontalRate,
        [HorizontalAverage] = (isnull(flc.HorizontalTest1, 0) + isnull(flc.HorizontalTest2, 0)  + isnull(flc.HorizontalTest3, 0)) / 3.0,
        [VerticalTest1] = flc.VerticalTest1,
        [VerticalTest2] = flc.VerticalTest2,
        [VerticalTest3] = flc.VerticalTest3,
        [VerticalRate] = flc.VerticalRate,
        [VerticalAverage] = (isnull(flc.VerticalTest1, 0) + isnull(flc.VerticalTest2, 0)  + isnull(flc.VerticalTest3, 0)) / 3.0,
        [Inspdate] = flc.Inspdate,
        [Inspector] = flc.Inspector,
        [Name] = (select Concat(Name, ' Ext.', ExtNo) from pass1 WITH(NOLOCK) where ID = flc.Inspector),
        [Remark] = flc.Remark,
        [LastUpdate] = Concat(LastUpdateName.val, ' - ', isnull(Format(flc.EditDate, 'yyyy/MM/dd HH:mm:ss'), Format(flc.AddDate, 'yyyy/MM/dd HH:mm:ss')))
from FIR_Laboratory_Heat flc with (nolock)
outer apply (select [val] = Name_Extno from View_ShowName where ID = iif(isnull(flc.EditName, '') = '', flc.AddName, flc.EditName)) LastUpdateName
where flc.ID = @ID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetFabricCrkShrkTestHeat_Detail, listPar);
        }


        #region Iron
        public FabricCrkShrkTestIron_Main GetFabricIronTest_Main(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrkShrkTestIron_Main = @"

select	[POID] = f.POID,
        [SEQ] = Concat(f.Seq1, ' ', f.Seq2),
        [ColorID] = pc.SpecValue,
        [ArriveQty] = f.ArriveQty,
        [WhseArrival] = r.WhseArrival,
        [ExportID] = r.ExportID,
        [Supp] = Concat(f.SuppID, s.AbbEn),
        [Iron] = fl.Iron,
        [IronDate] = fl.IronDate,
        [StyleID] = o.StyleID,
        [SCIRefno] = f.SCIRefno,
        [Name] = (select Name from pass1 WITH(NOLOCK) where ID = fl.IronInspector),
        [BrandID] = o.BrandID,
        [Refno] = f.Refno,
        [NonIron] = fl.NonIron,
        [DescDetail] = fab.DescDetail,
        [IronRemark] = fl.IronRemark,
        [IronEncode] = fl.IronEncode,
        [IronTestBeforePicture] =  (select top 1 IronTestBeforePicture from SciPMSFile_FIR_Laboratory fli WITH (NOLOCK) where fli.ID = fl.ID ),  
        [IronTestAfterPicture] = (select top 1 IronTestAfterPicture from SciPMSFile_FIR_Laboratory fli WITH (NOLOCK) where fli.ID = fl.ID),
        [ReportNo] = fl.ReportNo
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Supp s with (nolock) on s.ID = f.SuppID
left join Orders o with (nolock) on o.ID = f.POID
left join Fabric fab with (nolock) on fab.SCIRefno = f.SCIRefno
where f.ID = @ID
";

            IList<FabricCrkShrkTestIron_Main> listResult = ExecuteList<FabricCrkShrkTestIron_Main>(CommandType.Text, sqlGetFabricCrkShrkTestIron_Main, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult[0];
        }

        public List<FabricCrkShrkTestIron_Detail> GetFabricIronTest_Detail(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrkShrkTestIron_Detail = @"

select	[Roll] = flc.Roll,
        [Dyelot] = flc.Dyelot,
        [HorizontalOriginal] = flc.HorizontalOriginal,
        [VerticalOriginal] = flc.VerticalOriginal,
        [Result] = flc.Result,
        [HorizontalTest1] = flc.HorizontalTest1,
        [HorizontalTest2] = flc.HorizontalTest2,
        [HorizontalTest3] = flc.HorizontalTest3,
        [HorizontalRate] = flc.HorizontalRate,
        [HorizontalAverage] = Cast(Round((isnull(flc.HorizontalTest1, 0) + isnull(flc.HorizontalTest2, 0)  + isnull(flc.HorizontalTest3, 0)) / 3.0, 2) as numeric(5, 2)),
        [VerticalTest1] = flc.VerticalTest1,
        [VerticalTest2] = flc.VerticalTest2,
        [VerticalTest3] = flc.VerticalTest3,
        [VerticalRate] = flc.VerticalRate,
        [VerticalAverage] = Cast(Round((isnull(flc.VerticalTest1, 0) + isnull(flc.VerticalTest2, 0)  + isnull(flc.VerticalTest3, 0)) / 3.0, 2) as numeric(5, 2)),
        [Inspdate] = flc.Inspdate,
        [Inspector] = flc.Inspector,
        [Name] = (select Concat(Name, 'Ext.', ExtNo) from pass1 WITH(NOLOCK) where ID = flc.Inspector),
        [Remark] = flc.Remark,
        [LastUpdate] = Concat(LastUpdateName.val, ' - ', isnull(Format(flc.EditDate, 'yyyy/MM/dd HH:mm:ss'), Format(flc.AddDate, 'yyyy/MM/dd HH:mm:ss')))
from FIR_Laboratory_Iron flc with (nolock)
outer apply (select [val] = Name_Extno from View_ShowName where ID = iif(isnull(flc.EditName, '') = '', flc.AddName, flc.EditName)) LastUpdateName
where flc.ID = @ID
";

            return ExecuteList<FabricCrkShrkTestIron_Detail>(CommandType.Text, sqlGetFabricCrkShrkTestIron_Detail, listPar).ToList();
        }

        public void UpdateFabricIronTestDetail(FabricCrkShrkTestIron_Result fabricCrkShrkTestIron_Result, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", fabricCrkShrkTestIron_Result.ID);
            listPar.Add("@IronRemark", fabricCrkShrkTestIron_Result.Iron_Main.IronRemark ?? string.Empty);
            listPar.Add("@IronTestBeforePicture", fabricCrkShrkTestIron_Result.Iron_Main.IronTestBeforePicture);
            listPar.Add("@IronTestAfterPicture", fabricCrkShrkTestIron_Result.Iron_Main.IronTestAfterPicture);

            string sqlUpdateCrocking = @"
SET XACT_ABORT ON
-----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
update  FIR_Laboratory set  IronRemark = @IronRemark
where   ID = @ID 
;
if exists(
    select 1 from SciPMSFile_FIR_Laboratory where  ID = @ID 
)
begin
    update  SciPMSFile_FIR_Laboratory set IronTestBeforePicture = @IronTestBeforePicture,
                                IronTestAfterPicture = @IronTestAfterPicture
    where   ID = @ID 
end
else
begin
    insert into SciPMSFile_FIR_Laboratory (ID,IronTestBeforePicture,IronTestAfterPicture)
    values (@ID ,@IronTestBeforePicture ,@IronTestAfterPicture )
end
";


            List<FabricCrkShrkTestIron_Detail> oldIronData = GetFabricIronTest_Detail(fabricCrkShrkTestIron_Result.ID);

            List<FabricCrkShrkTestIron_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<FabricCrkShrkTestIron_Detail>(
                    fabricCrkShrkTestIron_Result.Iron_Detail,
                    oldIronData,
                    "Roll,Dyelot",
                    "HorizontalOriginal,VerticalOriginal,Result,HorizontalTest1,HorizontalTest2,HorizontalTest3,VerticalTest1,VerticalTest2,VerticalTest3,Inspdate,Inspector,Remark");

            string NewReportNo = GetID(fabricCrkShrkTestIron_Result.MDivisionID + "FT", "FIR_Laboratory", DateTime.Today, 2, "ReportNo");

            string sqlInsertDetail = @"
insert into FIR_Laboratory_Iron(
ID                   ,
Roll                 ,
Dyelot               ,
Inspdate             ,
Inspector            ,
Result               ,
Remark               ,
AddName              ,
AddDate              ,
HorizontalRate       ,
HorizontalOriginal   ,
HorizontalTest1      ,
HorizontalTest2      ,
HorizontalTest3      ,
VerticalRate         ,
VerticalOriginal     ,
VerticalTest1        ,
VerticalTest2        ,
VerticalTest3

)
values
(
@ID                   ,
@Roll                 ,
@Dyelot               ,
@Inspdate             ,
@Inspector            ,
@Result               ,
@Remark               ,
@AddName              ,
getDate()              ,
@HorizontalRate       ,
@HorizontalOriginal   ,
@HorizontalTest1      ,
@HorizontalTest2      ,
@HorizontalTest3      ,
@VerticalRate         ,
@VerticalOriginal     ,
@VerticalTest1        ,
@VerticalTest2        ,
@VerticalTest3
)
;
UPDATE FIR_Laboratory
SET ReportNo = @ReportNo
WHERE ReportNo = '' AND ID= @ID
";

            string sqlDeleteDetail = @"
delete  FIR_Laboratory_Iron
where   ID = @ID and
        Roll = @Roll and
        Dyelot = @Dyelot
";

            string sqlUpdateDetail = @"
update  FIR_Laboratory_Iron set Inspdate            = @Inspdate             ,
                                Inspector           = @Inspector            ,
                                Result              = @Result               ,
                                Remark              = @Remark               ,
                                EditName            = @EditName             ,
                                EditDate            = getDate()             ,
                                HorizontalRate      = @HorizontalRate       ,
                                HorizontalOriginal  = @HorizontalOriginal   ,
                                HorizontalTest1     = @HorizontalTest1      ,
                                HorizontalTest2     = @HorizontalTest2      ,
                                HorizontalTest3     = @HorizontalTest3      ,
                                VerticalRate        = @VerticalRate         ,
                                VerticalOriginal    = @VerticalOriginal     ,
                                VerticalTest1       = @VerticalTest1        ,
                                VerticalTest2       = @VerticalTest2        ,
                                VerticalTest3       = @VerticalTest3     
        where   ID = @ID and
                Roll = @Roll and
                Dyelot = @Dyelot
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdateCrocking, listPar);
                foreach (FabricCrkShrkTestIron_Detail detailItem in needUpdateDetailList)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();

                    switch (detailItem.StateType)
                    {
                        case DatabaseObject.Public.CompareStateType.Add:
                            listDetailPar.Add("@ID", fabricCrkShrkTestIron_Result.ID);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Inspdate", detailItem.Inspdate);
                            listDetailPar.Add("@Inspector", detailItem.Inspector);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@AddName", userID);
                            listDetailPar.Add("@HorizontalRate", detailItem.HorizontalRate);
                            listDetailPar.Add("@HorizontalOriginal", detailItem.HorizontalOriginal);
                            listDetailPar.Add("@HorizontalTest1", detailItem.HorizontalTest1);
                            listDetailPar.Add("@HorizontalTest2", detailItem.HorizontalTest2);
                            listDetailPar.Add("@HorizontalTest3", detailItem.HorizontalTest3);
                            listDetailPar.Add("@VerticalRate", detailItem.VerticalRate);
                            listDetailPar.Add("@VerticalOriginal", detailItem.VerticalOriginal);
                            listDetailPar.Add("@VerticalTest1", detailItem.VerticalTest1);
                            listDetailPar.Add("@VerticalTest2", detailItem.VerticalTest2);
                            listDetailPar.Add("@VerticalTest3", detailItem.VerticalTest3);
                            listDetailPar.Add("@ReportNo", NewReportNo);

                            ExecuteNonQuery(CommandType.Text, sqlInsertDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Edit:
                            listDetailPar.Add("@ID", fabricCrkShrkTestIron_Result.ID);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Inspdate", detailItem.Inspdate);
                            listDetailPar.Add("@Inspector", detailItem.Inspector);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@EditName", userID);
                            listDetailPar.Add("@HorizontalRate", detailItem.HorizontalRate);
                            listDetailPar.Add("@HorizontalOriginal", detailItem.HorizontalOriginal);
                            listDetailPar.Add("@HorizontalTest1", detailItem.HorizontalTest1);
                            listDetailPar.Add("@HorizontalTest2", detailItem.HorizontalTest2);
                            listDetailPar.Add("@HorizontalTest3", detailItem.HorizontalTest3);
                            listDetailPar.Add("@VerticalRate", detailItem.VerticalRate);
                            listDetailPar.Add("@VerticalOriginal", detailItem.VerticalOriginal);
                            listDetailPar.Add("@VerticalTest1", detailItem.VerticalTest1);
                            listDetailPar.Add("@VerticalTest2", detailItem.VerticalTest2);
                            listDetailPar.Add("@VerticalTest3", detailItem.VerticalTest3);

                            ExecuteNonQuery(CommandType.Text, sqlUpdateDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Delete:
                            listDetailPar.Add("@ID", fabricCrkShrkTestIron_Result.ID);
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

                string UpdateInspPercent = $@"
DECLARE @POID as varchar(15)= (SELECT TOP 1  POID from FIR_Laboratory where ID = @ID)
exec UpdateInspPercent 'FIRLab', @POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }


        public void EncodeFabricIron(long ID, string testResult, DateTime? IronDate, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);
            listPar.Add("@testResult", testResult);
            listPar.Add("@IronDate", IronDate);
            listPar.Add("@userID", userID);

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory  set Iron = @testResult,
                            IronDate = @IronDate,
                            IronEncode  = 1,
                            IronInspector = @userID
     where  ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}

declare @POID varchar(13)

select @POID = POID from FIR_Laboratory WITH(NOLOCK) where ID = @ID

exec UpdateInspPercent 'FIRLab',@POID 
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listPar);
                transaction.Complete();
            }
        }


        public DataTable GetIronFailMailContentData(long ID)
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
        [Color] = pc.SpecValue,
        [Supplier] = Concat(f.SuppID, s.AbbEn),
        [Arrive Qty] = f.ArriveQty,
        [Iron Result] = fl.Iron,
        [Iron Last Test Date] = Format(fl.IronDate, 'yyyy/MM/dd'),
        [Iron Remark] = fl.IronRemark
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Supp s with (nolock) on s.ID = f.SuppID
left join Orders o with (nolock) on o.ID = f.POID
left join Fabric fab with (nolock) on fab.SCIRefno = f.SCIRefno
where f.ID = @ID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public void AmendFabricIron(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory  set Iron = '',
                            IronDate = null,
                            IronEncode = 0,
                            IronInspector = ''
     where  ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}

declare @POID varchar(13)

select @POID = POID from FIR_Laboratory WITH(NOLOCK) where ID = @ID

exec UpdateInspPercent 'FIRLab',@POID 
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listPar);
                transaction.Complete();
            }
        }

        public DataTable GetIronDetailForReport(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrkShrkTestIron_Detail = @"

select	[Roll] = flc.Roll,
        [Dyelot] = flc.Dyelot,
        [HorizontalOriginal] = flc.HorizontalOriginal,
        [VerticalOriginal] = flc.VerticalOriginal,
        [Result] = flc.Result,
        [HorizontalTest1] = flc.HorizontalTest1,
        [HorizontalTest2] = flc.HorizontalTest2,
        [HorizontalTest3] = flc.HorizontalTest3,
        [HorizontalRate] = flc.HorizontalRate,
        [HorizontalAverage] = (isnull(flc.HorizontalTest1, 0) + isnull(flc.HorizontalTest2, 0)  + isnull(flc.HorizontalTest3, 0)) / 3.0,
        [VerticalTest1] = flc.VerticalTest1,
        [VerticalTest2] = flc.VerticalTest2,
        [VerticalTest3] = flc.VerticalTest3,
        [VerticalRate] = flc.VerticalRate,
        [VerticalAverage] = (isnull(flc.VerticalTest1, 0) + isnull(flc.VerticalTest2, 0)  + isnull(flc.VerticalTest3, 0)) / 3.0,
        [Inspdate] = flc.Inspdate,
        [Inspector] = flc.Inspector,
        [Name] = (select Concat(Name, ' Ext.', ExtNo) from pass1 WITH(NOLOCK) where ID = flc.Inspector),
        [Remark] = flc.Remark,
        [LastUpdate] = Concat(LastUpdateName.val, ' - ', isnull(Format(flc.EditDate, 'yyyy/MM/dd HH:mm:ss'), Format(flc.AddDate, 'yyyy/MM/dd HH:mm:ss')))
from FIR_Laboratory_Iron flc with (nolock)
outer apply (select [val] = Name_Extno from View_ShowName where ID = iif(isnull(flc.EditName, '') = '', flc.AddName, flc.EditName)) LastUpdateName
where flc.ID = @ID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetFabricCrkShrkTestIron_Detail, listPar);
        }

        #endregion

        public FabricCrkShrkTestWash_Main GetFabricWashTest_Main(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrkShrkTestWash_Main = @"

select	[POID] = f.POID,
        [SEQ] = Concat(f.Seq1, ' ', f.Seq2),
        [ColorID] = pc.SpecValue,
        [ArriveQty] = f.ArriveQty,
        [WhseArrival] = r.WhseArrival,
        [ExportID] = r.ExportID,
        [Supp] = Concat(f.SuppID, s.AbbEn),
        [Wash] = fl.Wash,
        [WashDate] = fl.WashDate,
        [StyleID] = o.StyleID,
        [SCIRefno] = f.SCIRefno,
        [Name] = (select Name from pass1 WITH(NOLOCK) where ID = fl.WashInspector),
        [BrandID] = o.BrandID,
        [Refno] = f.Refno,
        [NonWash] = fl.NonWash,
        [SkewnessOptionID] = fl.SkewnessOptionID,
        [DescDetail] = fab.DescDetail,
        [WashRemark] = fl.WashRemark,
        [WashEncode] = fl.WashEncode,
        [WashTestBeforePicture] = (select top 1 WashTestBeforePicture from SciPMSFile_FIR_Laboratory fli WITH(NOLOCK) where fli.ID = fl.ID),
        [WashTestAfterPicture] = (select top 1 WashTestAfterPicture from SciPMSFile_FIR_Laboratory fli WITH(NOLOCK) where fli.ID = fl.ID),
        [ReportNo] = fl.ReportNo
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Supp s with (nolock) on s.ID = f.SuppID
left join Orders o with (nolock) on o.ID = f.POID
left join Fabric fab with (nolock) on fab.SCIRefno = f.SCIRefno
where f.ID = @ID
";

            IList<FabricCrkShrkTestWash_Main> listResult = ExecuteList<FabricCrkShrkTestWash_Main>(CommandType.Text, sqlGetFabricCrkShrkTestWash_Main, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult[0];
        }

        public List<FabricCrkShrkTestWash_Detail> GetFabricWashTest_Detail(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrkShrkTestWash_Detail = @"

select	[Roll] = flc.Roll,
        [Dyelot] = flc.Dyelot,
        [HorizontalOriginal] = flc.HorizontalOriginal,
        [VerticalOriginal] = flc.VerticalOriginal,
        [Result] = flc.Result,
        [HorizontalTest1] = flc.HorizontalTest1,
        [HorizontalTest2] = flc.HorizontalTest2,
        [HorizontalTest3] = flc.HorizontalTest3,
        [HorizontalAverage] = Cast(Round((isnull(flc.HorizontalTest1, 0) + isnull(flc.HorizontalTest2, 0)  + isnull(flc.HorizontalTest3, 0)) / 3.0, 2) as numeric(5, 2)),
        [HorizontalRate] = flc.HorizontalRate,
        [VerticalTest1] = flc.VerticalTest1,
        [VerticalTest2] = flc.VerticalTest2,
        [VerticalTest3] = flc.VerticalTest3,
        [VerticalAverage] = Cast(Round((isnull(flc.VerticalTest1, 0) + isnull(flc.VerticalTest2, 0)  + isnull(flc.VerticalTest3, 0)) / 3.0, 2) as numeric(5, 2)),
        [VerticalRate] = flc.VerticalRate,
        [SkewnessTest1] = flc.SkewnessTest1,
        [SkewnessTest2] = flc.SkewnessTest2,
        [SkewnessTest3] = flc.SkewnessTest3,
        [SkewnessTest4] = flc.SkewnessTest4,
        [SkewnessRate] = flc.SkewnessRate,
        [Inspdate] = flc.Inspdate,
        [Inspector] = flc.Inspector,
        [Name] = (select Concat(Name, ' Ext.', ExtNo) from pass1 WITH(NOLOCK) where ID = flc.Inspector),
        [Remark] = flc.Remark,
        [LastUpdate] = Concat(LastUpdateName.val, ' - ', isnull(Format(flc.EditDate, 'yyyy/MM/dd HH:mm:ss'), Format(flc.AddDate, 'yyyy/MM/dd HH:mm:ss')))
from FIR_Laboratory_Wash flc with (nolock)
outer apply (select [val] = Name_Extno from View_ShowName where ID = iif(isnull(flc.EditName, '') = '', flc.AddName, flc.EditName)) LastUpdateName
where flc.ID = @ID
";

            return ExecuteList<FabricCrkShrkTestWash_Detail>(CommandType.Text, sqlGetFabricCrkShrkTestWash_Detail, listPar).ToList();
        }

        public void UpdateFabricWashTestDetail(FabricCrkShrkTestWash_Result fabricCrkShrkTestWash_Result, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", fabricCrkShrkTestWash_Result.ID);
            listPar.Add("@WashRemark", fabricCrkShrkTestWash_Result.Wash_Main.WashRemark);
            listPar.Add("@SkewnessOptionID", fabricCrkShrkTestWash_Result.Wash_Main.SkewnessOptionID);
            listPar.Add("@WashTestBeforePicture", fabricCrkShrkTestWash_Result.Wash_Main.WashTestBeforePicture);
            listPar.Add("@WashTestAfterPicture", fabricCrkShrkTestWash_Result.Wash_Main.WashTestAfterPicture);

            string sqlUpdateCrocking = @"
SET XACT_ABORT ON
-----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
update  FIR_Laboratory set  WashRemark = @WashRemark, 
                            SkewnessOptionID = @SkewnessOptionID
where   ID = @ID 
;
if exists(
    select 1 from SciPMSFile_FIR_Laboratory where  ID = @ID 
)
begin
    update  SciPMSFile_FIR_Laboratory set WashTestBeforePicture = @WashTestBeforePicture,
                                WashTestAfterPicture = @WashTestAfterPicture
    where   ID = @ID 
end
else
begin
    insert into SciPMSFile_FIR_Laboratory (ID,WashTestBeforePicture,WashTestAfterPicture)
    values (@ID ,@WashTestBeforePicture ,@WashTestAfterPicture )
end
";


            List<FabricCrkShrkTestWash_Detail> oldWashData = GetFabricWashTest_Detail(fabricCrkShrkTestWash_Result.ID);

            List<FabricCrkShrkTestWash_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<FabricCrkShrkTestWash_Detail>(
                    fabricCrkShrkTestWash_Result.Wash_Detail,
                    oldWashData,
                    "Roll,Dyelot",
                    "HorizontalOriginal,VerticalOriginal,Result,HorizontalTest1,HorizontalTest2,HorizontalTest3,VerticalTest1,VerticalTest2,VerticalTest3,SkewnessTest1,SkewnessTest2,SkewnessTest3,SkewnessTest4,Inspdate,Inspector,Remark");

            string NewReportNo = GetID(fabricCrkShrkTestWash_Result.MDivisionID + "FT", "FIR_Laboratory", DateTime.Today, 2, "ReportNo");

            string sqlInsertDetail = @"
insert into FIR_Laboratory_Wash(
ID                   ,
Roll                 ,
Dyelot               ,
Inspdate             ,
Inspector            ,
Result               ,
Remark               ,
AddName              ,
AddDate              ,
HorizontalRate       ,
HorizontalOriginal   ,
HorizontalTest1      ,
HorizontalTest2      ,
HorizontalTest3      ,
VerticalRate         ,
VerticalOriginal     ,
VerticalTest1        ,
VerticalTest2        ,
VerticalTest3       ,
SkewnessTest1       ,
SkewnessTest2       ,
SkewnessTest3       ,
SkewnessTest4       ,
SkewnessRate
)
values
(
@ID                   ,
@Roll                 ,
@Dyelot               ,
@Inspdate             ,
@Inspector            ,
@Result               ,
@Remark               ,
@AddName              ,
getDate()              ,
@HorizontalRate       ,
@HorizontalOriginal   ,
@HorizontalTest1      ,
@HorizontalTest2      ,
@HorizontalTest3      ,
@VerticalRate         ,
@VerticalOriginal     ,
@VerticalTest1        ,
@VerticalTest2        ,
@VerticalTest3      ,
@SkewnessTest1       ,
@SkewnessTest2       ,
@SkewnessTest3       ,
@SkewnessTest4       ,
@SkewnessRate
)
;
UPDATE FIR_Laboratory
SET ReportNo = @ReportNo
WHERE ReportNo = '' AND ID= @ID
";

            string sqlDeleteDetail = @"
delete  FIR_Laboratory_Wash
where   ID = @ID and
        Roll = @Roll and
        Dyelot = @Dyelot
";

            string sqlUpdateDetail = @"
update  FIR_Laboratory_Wash set Inspdate            = @Inspdate             ,
                                Inspector           = @Inspector            ,
                                Result              = @Result               ,
                                Remark              = @Remark               ,
                                EditName            = @EditName             ,
                                EditDate            = getDate()             ,
                                HorizontalRate      = @HorizontalRate       ,
                                HorizontalOriginal  = @HorizontalOriginal   ,
                                HorizontalTest1     = @HorizontalTest1      ,
                                HorizontalTest2     = @HorizontalTest2      ,
                                HorizontalTest3     = @HorizontalTest3      ,
                                VerticalRate        = @VerticalRate         ,
                                VerticalOriginal    = @VerticalOriginal     ,
                                VerticalTest1       = @VerticalTest1        ,
                                VerticalTest2       = @VerticalTest2        ,
                                VerticalTest3       = @VerticalTest3         ,
                                SkewnessTest1       = @SkewnessTest1,
                                SkewnessTest2       = @SkewnessTest2,
                                SkewnessTest3       = @SkewnessTest3,
                                SkewnessTest4       = @SkewnessTest4,
                                SkewnessRate        = @SkewnessRate
        where   ID = @ID and
                Roll = @Roll and
                Dyelot = @Dyelot
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdateCrocking, listPar);
                foreach (FabricCrkShrkTestWash_Detail detailItem in needUpdateDetailList)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();

                    switch (detailItem.StateType)
                    {
                        case DatabaseObject.Public.CompareStateType.Add:
                            listDetailPar.Add("@ID", fabricCrkShrkTestWash_Result.ID);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Inspdate", detailItem.Inspdate);
                            listDetailPar.Add("@Inspector", detailItem.Inspector);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@AddName", userID);
                            listDetailPar.Add("@HorizontalRate", detailItem.HorizontalRate);
                            listDetailPar.Add("@HorizontalOriginal", detailItem.HorizontalOriginal);
                            listDetailPar.Add("@HorizontalTest1", detailItem.HorizontalTest1);
                            listDetailPar.Add("@HorizontalTest2", detailItem.HorizontalTest2);
                            listDetailPar.Add("@HorizontalTest3", detailItem.HorizontalTest3);
                            listDetailPar.Add("@VerticalRate", detailItem.VerticalRate);
                            listDetailPar.Add("@VerticalOriginal", detailItem.VerticalOriginal);
                            listDetailPar.Add("@VerticalTest1", detailItem.VerticalTest1);
                            listDetailPar.Add("@VerticalTest2", detailItem.VerticalTest2);
                            listDetailPar.Add("@VerticalTest3", detailItem.VerticalTest3);
                            listDetailPar.Add("@SkewnessTest1", detailItem.SkewnessTest1);
                            listDetailPar.Add("@SkewnessTest2", detailItem.SkewnessTest2);
                            listDetailPar.Add("@SkewnessTest3", detailItem.SkewnessTest3);
                            listDetailPar.Add("@SkewnessTest4", detailItem.SkewnessTest4);
                            listDetailPar.Add("@SkewnessRate", detailItem.SkewnessRate);
                            listDetailPar.Add("@ReportNo", NewReportNo);

                            ExecuteNonQuery(CommandType.Text, sqlInsertDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Edit:
                            listDetailPar.Add("@ID", fabricCrkShrkTestWash_Result.ID);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Inspdate", detailItem.Inspdate);
                            listDetailPar.Add("@Inspector", detailItem.Inspector);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@EditName", userID);
                            listDetailPar.Add("@HorizontalRate", detailItem.HorizontalRate);
                            listDetailPar.Add("@HorizontalOriginal", detailItem.HorizontalOriginal);
                            listDetailPar.Add("@HorizontalTest1", detailItem.HorizontalTest1);
                            listDetailPar.Add("@HorizontalTest2", detailItem.HorizontalTest2);
                            listDetailPar.Add("@HorizontalTest3", detailItem.HorizontalTest3);
                            listDetailPar.Add("@VerticalRate", detailItem.VerticalRate);
                            listDetailPar.Add("@VerticalOriginal", detailItem.VerticalOriginal);
                            listDetailPar.Add("@VerticalTest1", detailItem.VerticalTest1);
                            listDetailPar.Add("@VerticalTest2", detailItem.VerticalTest2);
                            listDetailPar.Add("@VerticalTest3", detailItem.VerticalTest3);
                            listDetailPar.Add("@SkewnessTest1", detailItem.SkewnessTest1);
                            listDetailPar.Add("@SkewnessTest2", detailItem.SkewnessTest2);
                            listDetailPar.Add("@SkewnessTest3", detailItem.SkewnessTest3);
                            listDetailPar.Add("@SkewnessTest4", detailItem.SkewnessTest4);
                            listDetailPar.Add("@SkewnessRate", detailItem.SkewnessRate);

                            ExecuteNonQuery(CommandType.Text, sqlUpdateDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Delete:
                            listDetailPar.Add("@ID", fabricCrkShrkTestWash_Result.ID);
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


                string UpdateInspPercent = $@"
DECLARE @POID as varchar(15)= (SELECT TOP 1  POID from FIR_Laboratory where ID = @ID)
exec UpdateInspPercent 'FIRLab', @POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }

        public void EncodeFabricWash(long ID, string testResult, DateTime? WashDate, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);
            listPar.Add("@testResult", testResult);
            listPar.Add("@WashDate", WashDate);
            listPar.Add("@userID", userID);

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory  set Wash = @testResult,
                            WashDate = @WashDate,
                            WashEncode  = 1,
                            WashInspector = @userID
     where  ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}

declare @POID varchar(13)

select @POID = POID from FIR_Laboratory WITH(NOLOCK) where ID = @ID

exec UpdateInspPercent 'FIRLab',@POID 
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listPar);
                transaction.Complete();
            }
        }

        public DataTable GetWashFailMailContentData(long ID)
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
        [Color] = pc.SpecValue,
        [Supplier] = Concat(f.SuppID, s.AbbEn),
        [Arrive Qty] = f.ArriveQty,
        [Wash Result] = fl.Wash,
        [Wash Last Test Date] = Format(fl.WashDate, 'yyyy/MM/dd'),
        [Wash Remark] = fl.WashRemark
from FIR f with (nolock)
left join FIR_Laboratory fl WITH (NOLOCK) on f.ID = fl.ID
left join Receiving r WITH (NOLOCK) on r.id = f.receivingid
left join Po_Supp_Detail psd with (nolock) on psd.ID = f.POID and psd.Seq1 = f.Seq1 and psd.Seq2 = f.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Supp s with (nolock) on s.ID = f.SuppID
left join Orders o with (nolock) on o.ID = f.POID
left join Fabric fab with (nolock) on fab.SCIRefno = f.SCIRefno
where f.ID = @ID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public void AmendFabricWash(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlUpdateFIR_Laboratory = $@"
update  FIR_Laboratory  set Wash = '',
                            WashDate = null,
                            WashEncode = 0,
                            WashInspector = ''
     where  ID = @ID

{FIR_Laboratory_Utility.UpdateResultSql}

declare @POID varchar(13)

select @POID = POID from FIR_Laboratory WITH(NOLOCK) where ID = @ID

exec UpdateInspPercent 'FIRLab',@POID 
";

            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdateFIR_Laboratory, listPar);
                transaction.Complete();
            }
        }

        public DataTable GetWashDetailForReport(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", ID);

            string sqlGetFabricCrkShrkTestWash_Detail = @"

select	[Roll] = flc.Roll,
        [Dyelot] = flc.Dyelot,
        [HorizontalOriginal] = flc.HorizontalOriginal,
        [VerticalOriginal] = flc.VerticalOriginal,
        [Result] = flc.Result,
        [HorizontalTest1] = flc.HorizontalTest1,
        [HorizontalTest2] = flc.HorizontalTest2,
        [HorizontalTest3] = flc.HorizontalTest3,
        [HorizontalRate] = flc.HorizontalRate,
        [HorizontalAverage] = (isnull(flc.HorizontalTest1, 0) + isnull(flc.HorizontalTest2, 0)  + isnull(flc.HorizontalTest3, 0)) / 3.0,
        [VerticalTest1] = flc.VerticalTest1,
        [VerticalTest2] = flc.VerticalTest2,
        [VerticalTest3] = flc.VerticalTest3,
        [VerticalRate] = flc.VerticalRate,
        [VerticalAverage] = (isnull(flc.VerticalTest1, 0) + isnull(flc.VerticalTest2, 0)  + isnull(flc.VerticalTest3, 0)) / 3.0,
        [SkewnessTest1] = flc.SkewnessTest1,
        [SkewnessTest2] = flc.SkewnessTest2,
        [SkewnessTest3] = flc.SkewnessTest3,
        [SkewnessTest4] = flc.SkewnessTest4,
        [SkewnessRate] = flc.SkewnessRate,
        [Inspdate] = flc.Inspdate,
        [Inspector] = flc.Inspector,
        [Name] = (select Concat(Name, ' Ext.', ExtNo) from pass1 WITH(NOLOCK) where ID = flc.Inspector),
        [Remark] = flc.Remark,
        [LastUpdate] = Concat(LastUpdateName.val, ' - ', isnull(Format(flc.EditDate, 'yyyy/MM/dd HH:mm:ss'), Format(flc.AddDate, 'yyyy/MM/dd HH:mm:ss')))
from FIR_Laboratory_Wash flc with (nolock)
outer apply (select [val] = Name_Extno from View_ShowName where ID = iif(isnull(flc.EditName, '') = '', flc.AddName, flc.EditName)) LastUpdateName
where flc.ID = @ID
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetFabricCrkShrkTestWash_Detail, listPar);
        }

        #endregion

    }
}
