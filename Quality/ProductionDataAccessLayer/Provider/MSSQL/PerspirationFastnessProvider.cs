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
    public class PerspirationFastnessProvider : SQLDAL
    {
        #region 底層連線
        public PerspirationFastnessProvider(string ConString) : base(ConString) { }
        public PerspirationFastnessProvider(SQLDataTransaction tra) : base(tra) { }

        #endregion

        public PerspirationFastness_Detail_Result GetPerspirationFastness_Detail(string poID, string TestNo)
        {
            PerspirationFastness_Detail_Result PerspirationFastness_Detail_Result = new PerspirationFastness_Detail_Result();

            if (string.IsNullOrEmpty(TestNo))
            {
                PerspirationFastness_Detail_Result.Main.Status = "";
                PerspirationFastness_Detail_Result.Main.POID = poID;
                return PerspirationFastness_Detail_Result;
            }

            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", poID);
            listPar.Add("@TestNo", decimal.Parse(TestNo));

            string sqlGetPerspirationFastness_Detail = @"

select	[TestNo] = cast(o.TestNo as varchar),
        [POID] = o.POID,
        o.ID, 
        o.ReportNo, 
		[InspDate] = o.InspDate,
		[Article] = o.Article,
		[Inspector] = o.Inspector,
        [InspectorName] = pass1Inspector.Name,
        [Result] = o.Result,
		[Remark] = o.Remark,
		[Status] = o.Status,
        Temperature = Cast( o.Temperature as varchar),
        Time = Cast(  o.Time as varchar),
        o.MetalContent
        ,TestBeforePicture = (select top 1 TestBeforePicture from SciPMSFile_PerspirationFastness oi WITH(NOLOCK) where  o.ID = oi.ID)
        ,TestAfterPicture = (select top 1 TestAfterPicture from SciPMSFile_PerspirationFastness oi WITH(NOLOCK) where o.ID = oi.ID)
from PerspirationFastness o with (nolock)
left join pass1 pass1Inspector WITH(NOLOCK) on o.Inspector = pass1Inspector.ID
where o.POID = @POID and o.TestNo = @TestNo
";

            IList<PerspirationFastness_Detail_Main> listPerspirationFastness_Detail = ExecuteList<PerspirationFastness_Detail_Main>(CommandType.Text, sqlGetPerspirationFastness_Detail, listPar);

            if (listPerspirationFastness_Detail.Count == 0)
            {
                throw new Exception($"TestNo<{TestNo}> data not found");
            }

            PerspirationFastness_Detail_Result.Main = listPerspirationFastness_Detail[0];

            string sqlGetDetails = @"
select	[SubmitDate] = od.SubmitDate,
        [PerspirationFastnessGroup] = od.PerspirationFastnessGroup,
        [SEQ] = Concat(od.Seq1, '-', od.Seq2),
        [Roll] = od.Roll,
        [Dyelot] = od.Dyelot,
        [Refno] = psd.Refno,
        [SCIRefno] = psd.SCIRefno,
        [ColorID] = pc.SpecValue,
        [Result] = od.Result
        ,od.AlkalineChangeScale
        ,od.AlkalineAcetateScale
        ,od.AlkalineCottonScale
        ,od.AlkalineNylonScale
        ,od.AlkalinePolyesterScale
        ,od.AlkalineAcrylicScale
        ,od.AlkalineWoolScale
        ,od.AlkalineResultChange
        ,od.AlkalineResultAcetate
        ,od.AlkalineResultCotton
        ,od.AlkalineResultNylon
        ,od.AlkalineResultPolyester
        ,od.AlkalineResultAcrylic
        ,od.AlkalineResultWool

        ,od.AcidChangeScale
        ,od.AcidAcetateScale
        ,od.AcidCottonScale
        ,od.AcidNylonScale
        ,od.AcidPolyesterScale
        ,od.AcidAcrylicScale
        ,od.AcidWoolScale
        ,od.AcidResultChange
        ,od.AcidResultAcetate
        ,od.AcidResultCotton
        ,od.AcidResultNylon
        ,od.AcidResultPolyester
        ,od.AcidResultAcrylic
        ,od.AcidResultWool,
        [Remark] = od.Remark,
        [LastUpdate] = Concat(od.EditName, '-', pass1EditName.Name, ' ', pass1EditName.Extno),
        [Temperature] = cast(o.Temperature as varchar),
        [Time] = cast(o.Time as varchar)
from PerspirationFastness_Detail od with (nolock)
inner join PerspirationFastness o with (nolock) on o.ID = od.ID
left join PO_Supp_Detail psd with (nolock) on o.POID = psd.ID and od.SEQ1 = psd.SEQ1 and od.SEQ2 = psd.SEQ2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join pass1 pass1EditName on od.EditName = pass1EditName.ID
where   o.POID = @POID and o.TestNo = @TestNo
";

            PerspirationFastness_Detail_Result.Details = ExecuteList<PerspirationFastness_Detail_Detail>(CommandType.Text, sqlGetDetails, listPar).ToList();

            return PerspirationFastness_Detail_Result;
        }

        public PerspirationFastness_Result GetPerspirationFastness_Main(string POID)
        {
            PerspirationFastness_Result PerspirationFastness_Result = new PerspirationFastness_Result();

            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", POID);

            string sqlGetPerspirationFastness_Main = @"
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
		[CompletionDate] = iif(p.LabPerspirationFastnessPercent >= 100, (select max(InspDate) from PerspirationFastness WITH(NOLOCK) where POID = p.ID), null),
		[LabPerspirationFastnessPercent] = p.LabPerspirationFastnessPercent,
		[Remark] = p.PerspirationFastnessLaboratoryRemark,
		[CreateBy] = Concat(p.AddName, '-', pass1AddName.Name, ' ', Format(p.AddDate, 'yyyy/MM/dd HH:mm:ss')),
		[EditBy] = Concat(p.EditName, '-', pass1EditName.Name, ' ', Format(p.EditDate, 'yyyy/MM/dd HH:mm:ss'))
from PO p with (nolock)
inner join Orders o with (nolock) on p.ID = o.ID
left join pass1 pass1AddName WITH(NOLOCK) on p.AddName = pass1AddName.ID
left join pass1 pass1EditName WITH(NOLOCK) on p.EditName = pass1EditName.ID
where p.id = @POID
";

            IList<PerspirationFastness_Main> listPerspirationFastness_Main = ExecuteList<PerspirationFastness_Main>(CommandType.Text, sqlGetPerspirationFastness_Main, listPar);

            if (listPerspirationFastness_Main.Count == 0)
            {
                throw new Exception($"PO<{POID}> data not found");
            }

            PerspirationFastness_Result.Main = listPerspirationFastness_Main[0];

            string sqlGetDetails = @"
select	[TestNo] = cast(o.TestNo as varchar),
        o.ReportNo ,
		[InspDate] = o.InspDate,
		[Article] = o.Article,
		[Result] = o.Result,
		[Inspector] = o.Inspector,
		[Remark] = o.Remark,
		[LastUpdate] = iif(	o.EditDate is null, 
							Concat(pass1AddName.Name, ' ', Format(o.AddDate, 'yyyy/MM/dd HH:mm:ss')),
							Concat(pass1EditName.Name, ' ', Format(o.EditDate, 'yyyy/MM/dd HH:mm:ss'))),
		[Status] = o.Status
from PerspirationFastness o with (nolock)
left join pass1 pass1AddName WITH(NOLOCK) on o.AddName = pass1AddName.ID
left join pass1 pass1EditName WITH(NOLOCK) on o.EditName = pass1EditName.ID
where o.POID = @POID
";

            PerspirationFastness_Result.Details = ExecuteList<PerspirationFastness_Detail>(CommandType.Text, sqlGetDetails, listPar).ToList();

            return PerspirationFastness_Result;
        }

        public void EditPerspirationFastnessDetail(PerspirationFastness_Detail_Result PerspirationFastness_Detail_Result, string userID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", PerspirationFastness_Detail_Result.Main.POID);
            listPar.Add("@TestNo", PerspirationFastness_Detail_Result.Main.TestNo);
            listPar.Add("@InspDate", PerspirationFastness_Detail_Result.Main.InspDate);
            listPar.Add("@Article", PerspirationFastness_Detail_Result.Main.Article);
            listPar.Add("@Inspector", PerspirationFastness_Detail_Result.Main.Inspector);
            listPar.Add("@Remark", PerspirationFastness_Detail_Result.Main.Remark ?? "");
            listPar.Add("@editName", userID);
            listPar.Add("@MetalContent", PerspirationFastness_Detail_Result.Main.MetalContent ?? "");
            listPar.Add("@TestBeforePicture", PerspirationFastness_Detail_Result.Main.TestBeforePicture);
            listPar.Add("@TestAfterPicture", PerspirationFastness_Detail_Result.Main.TestAfterPicture);
            listPar.Add("@Temperature", DbType.Int32, PerspirationFastness_Detail_Result.Main.Temperature);
            listPar.Add("@Time", DbType.Int32, PerspirationFastness_Detail_Result.Main.Time);

            string sqlUpdatePerspirationFastness = @"
SET XACT_ABORT ON
-----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
update  PerspirationFastness set    InspDate = @InspDate,
                    Article = @Article,
                    Inspector = @Inspector,
                    Temperature = @Temperature,
                    Time = @Time,
                    MetalContent = @MetalContent,
                    Remark = @Remark,
                    EditName = @editName,
                    EditDate = getdate()
where   POID = @POID and TestNo = @TestNo


select  [PerspirationFastnessID] = ID
from    PerspirationFastness WITH(NOLOCK)
where   POID = @POID and TestNo = @TestNo

----判斷ExtendServer有沒有缺資料
IF EXISTS(
    select 1 from PerspirationFastness a
    where NOT  EXISTS (
        select   ID
        from  SciPMSFile_PerspirationFastness b WITH(NOLOCK)
		where  a.ID = b.ID
    )
    AND   POID = @POID and TestNo = @TestNo
)
BEGIN
    insert into SciPMSFile_PerspirationFastness(ID, TestBeforePicture, TestAfterPicture)
    values
    (
            (
                select TOP 1 ID from PerspirationFastness a  WITH(NOLOCK)
                where NOT  EXISTS (
                    select   ID
                    from    SciPMSFile_PerspirationFastness b WITH(NOLOCK)
		            where  a.ID = b.ID
                )
                AND POID = @POID and TestNo = @TestNo
            )
            , @TestBeforePicture, @TestAfterPicture
    )
END
ELSE
BEGIN
    update  SciPMSFile_PerspirationFastness set  
                        TestBeforePicture = @TestBeforePicture,
                        TestAfterPicture = @TestAfterPicture
    where  ID IN (
        select   ID
        from    PerspirationFastness WITH(NOLOCK)
        where   POID = @POID and TestNo = @TestNo
    )
END


";


            PerspirationFastness_Detail_Result oldPerspirationFastnessData = GetPerspirationFastness_Detail(PerspirationFastness_Detail_Result.Main.POID, PerspirationFastness_Detail_Result.Main.TestNo);

            List<PerspirationFastness_Detail_Detail> needUpdateDetailList =
                PublicClass.CompareListValue<PerspirationFastness_Detail_Detail>(
                    PerspirationFastness_Detail_Result.Details,
                    oldPerspirationFastnessData.Details,
                    "PerspirationFastnessGroup,SEQ",
                    "Result,SubmitDate,Roll,AlkalineChangeScale,AlkalineResultChange" +
                    ",AlkalineAcetateScale,AlkalineResultAcetate,AlkalineCottonScale,AlkalineResultCotton" +
                    ",AlkalineNylonScale,AlkalineResultNylon,AlkalinePolyesterScale,AlkalineResultPolyester" +
                    ",AlkalineAcrylicScale,AlkalineResultAcrylic" +
                    ",AlkalineWoolScale,AlkalineResultWool" +

                    ",AcidChangeScale,AcidResultChange,AcidAcetateScale,AcidResultAcetate,AcidCottonScale,AcidResultCotton" +
                    ",AcidNylonScale,AcidResultNylon,AcidPolyesterScale,AcidResultPolyester" +
                    ",AcidAcrylicScale,AcidResultAcrylic" +
                    ",AcidWoolScale,AcidResultWool,Remark");



            string sqlInsertPerspirationFastnessDetail = @"
insert into PerspirationFastness_Detail(
ID             ,
PerspirationFastnessGroup      ,
SEQ1           ,
SEQ2           ,
Roll           ,
Dyelot         ,
Result         

,AlkalineChangeScale    
,AlkalineResultChange   
,AlkalineAcetateScale
,AlkalineResultAcetate
,AlkalineCottonScale
,AlkalineResultCotton
,AlkalineNylonScale
,AlkalineResultNylon
,AlkalinePolyesterScale
,AlkalineResultPolyester
,AlkalineAcrylicScale
,AlkalineResultAcrylic
,AlkalineWoolScale
,AlkalineResultWool


,AcidChangeScale    
,AcidResultChange   
,AcidAcetateScale
,AcidResultAcetate
,AcidCottonScale
,AcidResultCotton
,AcidNylonScale
,AcidResultNylon
,AcidPolyesterScale
,AcidResultPolyester
,AcidAcrylicScale
,AcidResultAcrylic
,AcidWoolScale
,AcidResultWool

,Remark         ,
AddName        ,
AddDate        ,
SubmitDate     
)
values
(
@ID             ,
@PerspirationFastnessGroup      ,
@SEQ1           ,
@SEQ2           ,
@Roll           ,
@Dyelot         ,
@Result         

,@AlkalineChangeScale    
,@AlkalineResultChange   
,@AlkalineAcetateScale
,@AlkalineResultAcetate
,@AlkalineCottonScale
,@AlkalineResultCotton
,@AlkalineNylonScale
,@AlkalineResultNylon
,@AlkalinePolyesterScale
,@AlkalineResultPolyester
,@AlkalineAcrylicScale
,@AlkalineResultAcrylic
,@AlkalineWoolScale
,@AlkalineResultWool

,@AcidChangeScale    
,@AcidResultChange   
,@AcidAcetateScale
,@AcidResultAcetate
,@AcidCottonScale
,@AcidResultCotton
,@AcidNylonScale
,@AcidResultNylon
,@AcidPolyesterScale
,@AcidResultPolyester
,@AcidAcrylicScale
,@AcidResultAcrylic
,@AcidWoolScale
,@AcidResultWool

,@Remark         ,
@AddName        ,
getdate()        ,
@SubmitDate     
)

";

            string sqlDeleteDetail = @"
delete  PerspirationFastness_Detail
where   ID = @ID and
        PerspirationFastnessGroup = @PerspirationFastnessGroup and
        SEQ1 = @SEQ1 and
        SEQ2 = @SEQ2

exec UpdateInspPercent 'LabPerspirationFastness',@POID
";

            string sqlUpdateDetail = @"
update  PerspirationFastness_Detail set Roll           =  @Roll         ,
                        Dyelot         =  @Dyelot       ,
                        Result         =  @Result       ,

                        AlkalineChangeScale    =  @AlkalineChangeScale  ,
                        AlkalineResultChange   =  @AlkalineResultChange ,
                        AlkalineAcetateScale    =  @AlkalineAcetateScale  ,
                        AlkalineResultAcetate   =  @AlkalineResultAcetate ,
                        AlkalineCottonScale    =  @AlkalineCottonScale  ,
                        AlkalineResultCotton   =  @AlkalineResultCotton ,
                        AlkalineNylonScale    =  @AlkalineNylonScale  ,
                        AlkalineResultNylon   =  @AlkalineResultNylon ,
                        AlkalinePolyesterScale    =  @AlkalinePolyesterScale  ,
                        AlkalineResultPolyester   =  @AlkalineResultPolyester ,
                        AlkalineAcrylicScale    =  @AlkalineAcrylicScale  ,
                        AlkalineResultAcrylic   =  @AlkalineResultAcrylic ,
                        AlkalineWoolScale    =  @AlkalineWoolScale  ,
                        AlkalineResultWool   =  @AlkalineResultWool ,

                        AcidChangeScale    =  @AcidChangeScale  ,
                        AcidResultChange   =  @AcidResultChange ,
                        AcidAcetateScale    =  @AcidAcetateScale  ,
                        AcidResultAcetate   =  @AcidResultAcetate ,
                        AcidCottonScale    =  @AcidCottonScale  ,
                        AcidResultCotton   =  @AcidResultCotton ,
                        AcidNylonScale    =  @AcidNylonScale  ,
                        AcidResultNylon   =  @AcidResultNylon ,
                        AcidPolyesterScale    =  @AcidPolyesterScale  ,
                        AcidResultPolyester   =  @AcidResultPolyester ,
                        AcidAcrylicScale    =  @AcidAcrylicScale  ,
                        AcidResultAcrylic   =  @AcidResultAcrylic ,
                        AcidWoolScale    =  @AcidWoolScale  ,
                        AcidResultWool   =  @AcidResultWool ,

                        Remark         =  @Remark       ,
                        SubmitDate     =  @SubmitDate   ,
                        EditName = @EditName,
                        EditDate = getdate()
        where   ID = @ID and
                PerspirationFastnessGroup = @PerspirationFastnessGroup and
                SEQ1 = @SEQ1 and
                SEQ2 = @SEQ2
";

            using (TransactionScope transaction = new TransactionScope())
            {
                DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdatePerspirationFastness, listPar);
                long PerspirationFastnessID = (long)dtResult.Rows[0]["PerspirationFastnessID"];
                foreach (PerspirationFastness_Detail_Detail detailItem in needUpdateDetailList)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();

                    switch (detailItem.StateType)
                    {
                        case DatabaseObject.Public.CompareStateType.Add:
                            listDetailPar.Add("@ID", PerspirationFastnessID);
                            listDetailPar.Add("@PerspirationFastnessGroup", detailItem.PerspirationFastnessGroup);
                            listDetailPar.Add("@SEQ1", detailItem.Seq1);
                            listDetailPar.Add("@SEQ2", detailItem.Seq2);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Result", detailItem.Result);

                            listDetailPar.Add("@AlkalineChangeScale", detailItem.AlkalineChangeScale);
                            listDetailPar.Add("@AlkalineResultChange", detailItem.AlkalineResultChange);

                            listDetailPar.Add("@AlkalineAcetateScale", detailItem.AlkalineAcetateScale);
                            listDetailPar.Add("@AlkalineResultAcetate", detailItem.AlkalineResultAcetate);

                            listDetailPar.Add("@AlkalineCottonScale", detailItem.AlkalineCottonScale);
                            listDetailPar.Add("@AlkalineResultCotton", detailItem.AlkalineResultCotton);

                            listDetailPar.Add("@AlkalineNylonScale", detailItem.AlkalineNylonScale);
                            listDetailPar.Add("@AlkalineResultNylon", detailItem.AlkalineResultNylon);

                            listDetailPar.Add("@AlkalinePolyesterScale", detailItem.AlkalinePolyesterScale);
                            listDetailPar.Add("@AlkalineResultPolyester", detailItem.AlkalineResultPolyester);

                            listDetailPar.Add("@AlkalineAcrylicScale", detailItem.AlkalineAcrylicScale);
                            listDetailPar.Add("@AlkalineResultAcrylic", detailItem.AlkalineResultAcrylic);

                            listDetailPar.Add("@AlkalineWoolScale", detailItem.AlkalineWoolScale);
                            listDetailPar.Add("@AlkalineResultWool", detailItem.AlkalineResultWool);


                            listDetailPar.Add("@AcidChangeScale", detailItem.AcidChangeScale);
                            listDetailPar.Add("@AcidResultChange", detailItem.AcidResultChange);

                            listDetailPar.Add("@AcidAcetateScale", detailItem.AcidAcetateScale);
                            listDetailPar.Add("@AcidResultAcetate", detailItem.AcidResultAcetate);

                            listDetailPar.Add("@AcidCottonScale", detailItem.AcidCottonScale);
                            listDetailPar.Add("@AcidResultCotton", detailItem.AcidResultCotton);

                            listDetailPar.Add("@AcidNylonScale", detailItem.AcidNylonScale);
                            listDetailPar.Add("@AcidResultNylon", detailItem.AcidResultNylon);

                            listDetailPar.Add("@AcidPolyesterScale", detailItem.AcidPolyesterScale);
                            listDetailPar.Add("@AcidResultPolyester", detailItem.AcidResultPolyester);

                            listDetailPar.Add("@AcidAcrylicScale", detailItem.AcidAcrylicScale);
                            listDetailPar.Add("@AcidResultAcrylic", detailItem.AcidResultAcrylic);

                            listDetailPar.Add("@AcidWoolScale", detailItem.AcidWoolScale);
                            listDetailPar.Add("@AcidResultWool", detailItem.AcidResultWool);


                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@AddName", userID);
                            listDetailPar.Add("@SubmitDate", detailItem.SubmitDate);

                            ExecuteNonQuery(CommandType.Text, sqlInsertPerspirationFastnessDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Edit:
                            listDetailPar.Add("@ID", PerspirationFastnessID);
                            listDetailPar.Add("@PerspirationFastnessGroup", detailItem.PerspirationFastnessGroup);
                            listDetailPar.Add("@SEQ1", detailItem.Seq1);
                            listDetailPar.Add("@SEQ2", detailItem.Seq2);
                            listDetailPar.Add("@Roll", detailItem.Roll);
                            listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                            listDetailPar.Add("@Result", detailItem.Result);
                            listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                            listDetailPar.Add("@EditName", userID);

                            listDetailPar.Add("@AlkalineChangeScale", detailItem.AlkalineChangeScale);
                            listDetailPar.Add("@AlkalineResultChange", detailItem.AlkalineResultChange);

                            listDetailPar.Add("@AlkalineAcetateScale", detailItem.AlkalineAcetateScale);
                            listDetailPar.Add("@AlkalineResultAcetate", detailItem.AlkalineResultAcetate);

                            listDetailPar.Add("@AlkalineCottonScale", detailItem.AlkalineCottonScale);
                            listDetailPar.Add("@AlkalineResultCotton", detailItem.AlkalineResultCotton);

                            listDetailPar.Add("@AlkalineNylonScale", detailItem.AlkalineNylonScale);
                            listDetailPar.Add("@AlkalineResultNylon", detailItem.AlkalineResultNylon);

                            listDetailPar.Add("@AlkalinePolyesterScale", detailItem.AlkalinePolyesterScale);
                            listDetailPar.Add("@AlkalineResultPolyester", detailItem.AlkalineResultPolyester);

                            listDetailPar.Add("@AlkalineAcrylicScale", detailItem.AlkalineAcrylicScale);
                            listDetailPar.Add("@AlkalineResultAcrylic", detailItem.AlkalineResultAcrylic);

                            listDetailPar.Add("@AlkalineWoolScale", detailItem.AlkalineWoolScale);
                            listDetailPar.Add("@AlkalineResultWool", detailItem.AlkalineResultWool);


                            listDetailPar.Add("@AcidChangeScale", detailItem.AcidChangeScale);
                            listDetailPar.Add("@AcidResultChange", detailItem.AcidResultChange);

                            listDetailPar.Add("@AcidAcetateScale", detailItem.AcidAcetateScale);
                            listDetailPar.Add("@AcidResultAcetate", detailItem.AcidResultAcetate);

                            listDetailPar.Add("@AcidCottonScale", detailItem.AcidCottonScale);
                            listDetailPar.Add("@AcidResultCotton", detailItem.AcidResultCotton);

                            listDetailPar.Add("@AcidNylonScale", detailItem.AcidNylonScale);
                            listDetailPar.Add("@AcidResultNylon", detailItem.AcidResultNylon);

                            listDetailPar.Add("@AcidPolyesterScale", detailItem.AcidPolyesterScale);
                            listDetailPar.Add("@AcidResultPolyester", detailItem.AcidResultPolyester);

                            listDetailPar.Add("@AcidAcrylicScale", detailItem.AcidAcrylicScale);
                            listDetailPar.Add("@AcidResultAcrylic", detailItem.AcidResultAcrylic);

                            listDetailPar.Add("@AcidWoolScale", detailItem.AcidWoolScale);
                            listDetailPar.Add("@AcidResultWool", detailItem.AcidResultWool);

                            listDetailPar.Add("@SubmitDate", detailItem.SubmitDate);

                            ExecuteNonQuery(CommandType.Text, sqlUpdateDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.Delete:
                            listDetailPar.Add("@ID", PerspirationFastnessID);
                            listDetailPar.Add("@PerspirationFastnessGroup", detailItem.PerspirationFastnessGroup);
                            listDetailPar.Add("@SEQ1", detailItem.Seq1);
                            listDetailPar.Add("@SEQ2", detailItem.Seq2);
                            listDetailPar.Add("@POID", PerspirationFastness_Detail_Result.Main.POID);

                            ExecuteNonQuery(CommandType.Text, sqlDeleteDetail, listDetailPar);
                            break;
                        case DatabaseObject.Public.CompareStateType.None:
                            break;
                        default:
                            break;
                    }
                }

                string UpdateInspPercent = "exec UpdateInspPercent 'LabPerspirationFastness',@POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }

        public void AddPerspirationFastnessDetail(PerspirationFastness_Detail_Result PerspirationFastness_Detail_Result, string userID, out string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", PerspirationFastness_Detail_Result.Main.POID);
            listPar.Add("@InspDate", PerspirationFastness_Detail_Result.Main.InspDate);
            listPar.Add("@Article", PerspirationFastness_Detail_Result.Main.Article);
            listPar.Add("@Inspector", PerspirationFastness_Detail_Result.Main.Inspector);
            listPar.Add("@Remark", PerspirationFastness_Detail_Result.Main.Remark ?? "");
            listPar.Add("@addName", userID);
            listPar.Add("@TestBeforePicture", PerspirationFastness_Detail_Result.Main.TestBeforePicture);
            listPar.Add("@TestAfterPicture", PerspirationFastness_Detail_Result.Main.TestAfterPicture);
            listPar.Add("@Remark", PerspirationFastness_Detail_Result.Main.Remark ?? "");
            listPar.Add("@MetalContent", PerspirationFastness_Detail_Result.Main.MetalContent);
            listPar.Add("@Temperature", DbType.Int32, PerspirationFastness_Detail_Result.Main.Temperature);
            listPar.Add("@Time", DbType.Int32, PerspirationFastness_Detail_Result.Main.Time);

            string NewReportNo = GetID(PerspirationFastness_Detail_Result.MDivisionID + "PF", "PerspirationFastness", DateTime.Today, 2, "ReportNo");
            listPar.Add("@ReportNo", NewReportNo);

            string sqlInsertPerspirationFastness = @"
SET XACT_ABORT ON

BEGIN DISTRIBUTED TRAN
declare @TestNo numeric(2,0)
DECLARE @PerspirationFastnessID table (ID bigint, TestNo numeric(2, 0))

select  @TestNo = isnull(Max(TestNo), 0) + 1
from    PerspirationFastness  WITH(NOLOCK)
where POID = @POID

----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
insert into PerspirationFastness(POID, TestNo, InspDate, Article, Status, Inspector, Temperature, MetalContent, Time , Remark, addName, addDate ,ReportNo)
        OUTPUT INSERTED.ID, INSERTED.TestNo into @PerspirationFastnessID
        values(@POID, @TestNo, @InspDate, @Article, 'New', @Inspector, @Temperature, @MetalContent, @Time, @Remark, @addName, getdate() ,@ReportNo)

select  [PerspirationFastnessID] = ID, TestNo
from @PerspirationFastnessID

insert into SciPMSFile_PerspirationFastness(ID, TestBeforePicture, TestAfterPicture)
        values(
(select ID from @PerspirationFastnessID) , @TestBeforePicture, @TestAfterPicture)

COMMIT TRAN
";

            string sqlInsertPerspirationFastnessDetail = @"
insert into PerspirationFastness_Detail(
ID             ,
PerspirationFastnessGroup      ,
SEQ1           ,
SEQ2           ,
Roll           ,
Dyelot         ,
Result         ,

AlkalineChangeScale    ,
AlkalineResultChange   ,
AlkalineAcetateScale    ,
AlkalineResultAcetate   ,
AlkalineCottonScale    ,
AlkalineResultCotton   ,
AlkalineNylonScale    ,
AlkalineResultNylon   ,
AlkalinePolyesterScale    ,
AlkalineResultPolyester   ,
AlkalineAcrylicScale    ,
AlkalineResultAcrylic   ,
AlkalineWoolScale    ,
AlkalineResultWool   ,

AcidChangeScale    ,
AcidResultChange   ,
AcidAcetateScale    ,
AcidResultAcetate   ,
AcidCottonScale    ,
AcidResultCotton   ,
AcidNylonScale    ,
AcidResultNylon   ,
AcidPolyesterScale    ,
AcidResultPolyester   ,
AcidAcrylicScale    ,
AcidResultAcrylic   ,
AcidWoolScale    ,
AcidResultWool   ,

Remark         ,
AddName        ,
AddDate        ,
SubmitDate     
)
values
(
@ID             ,
@PerspirationFastnessGroup      ,
@SEQ1           ,
@SEQ2           ,
@Roll           ,
@Dyelot         ,
@Result         ,

@AlkalineChangeScale    ,
@AlkalineResultChange   ,
@AlkalineAcetateScale    ,
@AlkalineResultAcetate   ,
@AlkalineCottonScale    ,
@AlkalineResultCotton   ,
@AlkalineNylonScale    ,
@AlkalineResultNylon   ,
@AlkalinePolyesterScale    ,
@AlkalineResultPolyester   ,
@AlkalineAcrylicScale    ,
@AlkalineResultAcrylic   ,
@AlkalineWoolScale    ,
@AlkalineResultWool   ,

@AcidChangeScale    ,
@AcidResultChange   ,
@AcidAcetateScale    ,
@AcidResultAcetate   ,
@AcidCottonScale    ,
@AcidResultCotton   ,
@AcidNylonScale    ,
@AcidResultNylon   ,
@AcidPolyesterScale    ,
@AcidResultPolyester   ,
@AcidAcrylicScale    ,
@AcidResultAcrylic   ,
@AcidWoolScale    ,
@AcidResultWool   ,
@Remark         ,
@AddName        ,
getdate()        ,
@SubmitDate     
)

";

            using (TransactionScope transaction = new TransactionScope())
            {
                DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlInsertPerspirationFastness, listPar);
                long PerspirationFastnessID = (long)dtResult.Rows[0]["PerspirationFastnessID"];
                TestNo = dtResult.Rows[0]["TestNo"].ToString();
                foreach (PerspirationFastness_Detail_Detail detailItem in PerspirationFastness_Detail_Result.Details)
                {
                    SQLParameterCollection listDetailPar = new SQLParameterCollection();
                    listDetailPar.Add("@ID", PerspirationFastnessID);
                    listDetailPar.Add("@PerspirationFastnessGroup", detailItem.PerspirationFastnessGroup);
                    listDetailPar.Add("@SEQ1", detailItem.Seq1);
                    listDetailPar.Add("@SEQ2", detailItem.Seq2);
                    listDetailPar.Add("@Roll", detailItem.Roll);
                    listDetailPar.Add("@Dyelot", detailItem.Dyelot);
                    listDetailPar.Add("@Result", detailItem.Result);
                    listDetailPar.Add("@AlkalineChangeScale", detailItem.AlkalineChangeScale);
                    listDetailPar.Add("@AlkalineResultChange", detailItem.AlkalineResultChange);

                    listDetailPar.Add("@AlkalineAcetateScale", detailItem.AlkalineAcetateScale);
                    listDetailPar.Add("@AlkalineResultAcetate", detailItem.AlkalineResultAcetate);

                    listDetailPar.Add("@AlkalineCottonScale", detailItem.AlkalineCottonScale);
                    listDetailPar.Add("@AlkalineResultCotton", detailItem.AlkalineResultCotton);

                    listDetailPar.Add("@AlkalineNylonScale", detailItem.AlkalineNylonScale);
                    listDetailPar.Add("@AlkalineResultNylon", detailItem.AlkalineResultNylon);

                    listDetailPar.Add("@AlkalinePolyesterScale", detailItem.AlkalinePolyesterScale);
                    listDetailPar.Add("@AlkalineResultPolyester", detailItem.AlkalineResultPolyester);

                    listDetailPar.Add("@AlkalineAcrylicScale", detailItem.AlkalineAcrylicScale);
                    listDetailPar.Add("@AlkalineResultAcrylic", detailItem.AlkalineResultAcrylic);

                    listDetailPar.Add("@AlkalineWoolScale", detailItem.AlkalineWoolScale);
                    listDetailPar.Add("@AlkalineResultWool", detailItem.AlkalineResultWool);


                    listDetailPar.Add("@AcidChangeScale", detailItem.AcidChangeScale);
                    listDetailPar.Add("@AcidResultChange", detailItem.AcidResultChange);

                    listDetailPar.Add("@AcidAcetateScale", detailItem.AcidAcetateScale);
                    listDetailPar.Add("@AcidResultAcetate", detailItem.AcidResultAcetate);

                    listDetailPar.Add("@AcidCottonScale", detailItem.AcidCottonScale);
                    listDetailPar.Add("@AcidResultCotton", detailItem.AcidResultCotton);

                    listDetailPar.Add("@AcidNylonScale", detailItem.AcidNylonScale);
                    listDetailPar.Add("@AcidResultNylon", detailItem.AcidResultNylon);

                    listDetailPar.Add("@AcidPolyesterScale", detailItem.AcidPolyesterScale);
                    listDetailPar.Add("@AcidResultPolyester", detailItem.AcidResultPolyester);

                    listDetailPar.Add("@AcidAcrylicScale", detailItem.AcidAcrylicScale);
                    listDetailPar.Add("@AcidResultAcrylic", detailItem.AcidResultAcrylic);

                    listDetailPar.Add("@AcidWoolScale", detailItem.AcidWoolScale);
                    listDetailPar.Add("@AcidResultWool", detailItem.AcidResultWool);

                    listDetailPar.Add("@Remark", detailItem.Remark ?? "");
                    listDetailPar.Add("@AddName", userID);
                    listDetailPar.Add("@SubmitDate", detailItem.SubmitDate);

                    ExecuteNonQuery(CommandType.Text, sqlInsertPerspirationFastnessDetail, listDetailPar);
                }

                string UpdateInspPercent = "exec UpdateInspPercent 'LabPerspirationFastness',@POID";
                ExecuteDataTableByServiceConn(CommandType.Text, UpdateInspPercent, listPar);

                transaction.Complete();
            }
        }

        public void SavePerspirationFastnessMain(PerspirationFastness_Main PerspirationFastness_Main)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@Remark", PerspirationFastness_Main.Remark ?? "");
            listPar.Add("@POID", PerspirationFastness_Main.POID);

            string sqlUpdatePerspirationFastnessMain = @"
update PO set PerspirationFastnessLaboratoryRemark = @Remark
where ID = @POID

exec UpdateInspPercent 'LabPerspirationFastness',@POID
";
            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlUpdatePerspirationFastnessMain, listPar);
                transaction.Complete();
            }

        }

        public void EncodePerspirationFastness(string poID, string TestNo, string result)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);
            listPar.Add("@result", result);

            string sqlUpdatePerspirationFastnessMain = @"
update PerspirationFastness set Status = 'Confirmed',
                Result = @result
where POID = @poID and TestNo = @TestNo

exec UpdateInspPercent 'LabPerspirationFastness',@poID
";
            ExecuteNonQuery(CommandType.Text, sqlUpdatePerspirationFastnessMain, listPar);
        }

        public void AmendPerspirationFastness(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlUpdatePerspirationFastnessMain = @"
update PerspirationFastness set Status = 'New',
                Result = ''
where POID = @poID and TestNo = @TestNo

exec UpdateInspPercent 'LabPerspirationFastness',@poID
";
            ExecuteNonQuery(CommandType.Text, sqlUpdatePerspirationFastnessMain, listPar);
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
from    PerspirationFastness ov with (nolock)
left join  Orders o with (nolock) on ov.POID = o.ID
where ov.POID = @poID and ov.TestNo = @TestNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public void DeletePerspirationFastness(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlDeletePerspirationFastness = @"
SET XACT_ABORT ON
delete  PerspirationFastness_Detail where ID = (select ID from PerspirationFastness where POID = @poID and TestNo = @TestNo)
delete  SciPMSFile_PerspirationFastness where ID = (select ID from PerspirationFastness where POID = @poID and TestNo = @TestNo)
delete  PerspirationFastness where POID = @poID and TestNo = @TestNo
exec UpdateInspPercent 'LabPerspirationFastness',@poID
";
            using (TransactionScope transaction = new TransactionScope())
            {
                ExecuteNonQuery(CommandType.Text, sqlDeletePerspirationFastness, listPar);
                transaction.Complete();
            }
        }

        public IList<PerspirationFastness_Excel> GetExcel(string ID)
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
        ,c.MetalContent
        ,c.Temperature
        ,c.Time
        ,cd.AlkalineChangeScale
        ,cd.AlkalineAcetateScale
        ,cd.AlkalineCottonScale
        ,cd.AlkalineNylonScale
        ,cd.AlkalinePolyesterScale
        ,cd.AlkalineAcrylicScale
        ,cd.AlkalineWoolScale
        ,cd.AlkalineResultChange
        ,cd.AlkalineResultAcetate
        ,cd.AlkalineResultCotton
        ,cd.AlkalineResultNylon
        ,cd.AlkalineResultPolyester
        ,cd.AlkalineResultAcrylic
        ,cd.AlkalineResultWool

        ,cd.AcidChangeScale
        ,cd.AcidAcetateScale
        ,cd.AcidCottonScale
        ,cd.AcidNylonScale
        ,cd.AcidPolyesterScale
        ,cd.AcidAcrylicScale
        ,cd.AcidWoolScale
        ,cd.AcidResultChange
        ,cd.AcidResultAcetate
        ,cd.AcidResultCotton
        ,cd.AcidResultNylon
        ,cd.AcidResultPolyester
        ,cd.AcidResultAcrylic
        ,cd.AcidResultWool

        ,cd.Remark
        ,c.Inspector
        ,TestBeforePicture = (select top 1 TestBeforePicture from SciPMSFile_PerspirationFastness pmsFile WITH(NOLOCK) where pmsFile.ID =  cd.ID)
        ,TestAfterPicture = (select top 1 TestAfterPicture from SciPMSFile_PerspirationFastness pmsFile WITH(NOLOCK) where pmsFile.ID =  cd.ID)
        ,c.ReportNo
from PerspirationFastness_Detail cd WITH(NOLOCK)
left join PerspirationFastness c WITH(NOLOCK) on c.ID =  cd.ID
left join Orders o WITH(NOLOCK) on o.ID=c.POID
left join PO_Supp_Detail psd WITH(NOLOCK) on c.POID = psd.ID and cd.SEQ1 = psd.SEQ1 and cd.SEQ2 = psd.SEQ2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Pass1 pEdit WITH(NOLOCK) on pEdit.ID = cd.EditName
left join pass1 pAdd WITH(NOLOCK) on pAdd.ID = cd.AddName
where cd.ID = @ID
order by cd.SubmitDate
";
            var detail = ExecuteList<PerspirationFastness_Excel>(CommandType.Text, sqlcmd, objParameter);

            return detail.Any() ? detail : new List<PerspirationFastness_Excel>();
        }
    }
}
