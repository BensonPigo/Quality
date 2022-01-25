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
		[InspDate] = o.InspDate,
		[Article] = o.Article,
		[Inspector] = o.Inspector,
        [InspectorName] = pass1Inspector.Name,
        [Result] = o.Result,
		[Remark] = o.Remark,
		[Status] = o.Status,
        Temperature = Cast( o.Temperature as varchar),
        Time = Cast(  o.Time as varchar),
        [TestBeforePicture] = oi.TestBeforePicture,
        [TestAfterPicture] = oi.TestAfterPicture
from PerspirationFastness o with (nolock)
LEFT JOIN [ExtendServer].PMSFile.dbo.PerspirationFastness oi with (nolock) ON o.ID = oi.ID
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
        [ColorID] = psd.ColorID,
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
                    Remark = @Remark,
                    EditName = @editName,
                    EditDate = getdate()
where   POID = @POID and TestNo = @TestNo


select  [PerspirationFastnessID] = ID
from    PerspirationFastness WITH(NOLOCK)
where   POID = @POID and TestNo = @TestNo

update  [ExtendServer].PMSFile.dbo.PerspirationFastness set  
                    TestBeforePicture = @TestBeforePicture,
                    TestAfterPicture = @TestAfterPicture
where  ID IN (
    select   ID
    from    PerspirationFastness WITH(NOLOCK)
    where   POID = @POID and TestNo = @TestNo
)

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
            listPar.Add("@Temperature", DbType.Int32, PerspirationFastness_Detail_Result.Main.Temperature);
            listPar.Add("@Time", DbType.Int32, PerspirationFastness_Detail_Result.Main.Time);

            string sqlInsertPerspirationFastness = @"
SET XACT_ABORT ON

BEGIN DISTRIBUTED TRAN
declare @TestNo numeric(2,0)
DECLARE @PerspirationFastnessID table (ID bigint, TestNo numeric(2, 0))

select  @TestNo = isnull(Max(TestNo), 0) + 1
from    PerspirationFastness  WITH(NOLOCK)
where POID = @POID

----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
insert into PerspirationFastness(POID, TestNo, InspDate, Article, Status, Inspector, Temperature, Time , Remark, addName, addDate)
        OUTPUT INSERTED.ID, INSERTED.TestNo into @PerspirationFastnessID
        values(@POID, @TestNo, @InspDate, @Article, 'New', @Inspector, @Temperature, @Time, @Remark, @addName, getdate())

select  [PerspirationFastnessID] = ID, TestNo
from @PerspirationFastnessID

insert into [ExtendServer].PMSFile.dbo.PerspirationFastness(ID, TestBeforePicture, TestAfterPicture)
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
--Temperature    ,
--Time
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
--@Temperature    ,
--@Time
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
        [Remark] = ov.Remark
from    PerspirationFastness ov with (nolock)
left join  Orders o with (nolock) on ov.POID = o.ID
where ov.POID = @poID and ov.TestNo = @TestNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public DataTable GetPerspirationFastnessDetailForExcel(string poID, string TestNo)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@poID", poID);
            listPar.Add("@TestNo", TestNo);

            string sqlGetData = @"
select	[SubmitDate] = od.SubmitDate,
        [PerspirationFastnessGroup] = od.PerspirationFastnessGroup,
        [SEQ] = Concat(od.Seq1, '-', od.Seq2),
        [Roll] = od.Roll,
        [Dyelot] = od.Dyelot,
        [Refno] = psd.Refno,
        [SCIRefno] = psd.SCIRefno,
        [ColorID] = psd.ColorID,
        [Result] = od.Result,

        [AlkalineChangeScale] = od.AlkalineChangeScale,
        [AlkalineResultChange] = od.AlkalineResultChange,
        [AlkalineAcetateScale] = od.AlkalineAcetateScale,
        [AlkalineResultAcetate] = od.AlkalineResultAcetate,
        [AlkalineCottonScale] = od.AlkalineCottonScale,
        [AlkalineResultCotton] = od.AlkalineResultCotton,
        [AlkalineNylonScale] = od.AlkalineNylonScale,
        [AlkalineResultNylon] = od.AlkalineResultNylon,
        [AlkalinePolyesterScale] = od.AlkalinePolyesterScale,
        [AlkalineResultPolyester] = od.AlkalineResultPolyester,
        [AlkalineAcrylicScale] = od.AlkalineAcrylicScale,
        [AlkalineResultAcrylic] = od.AlkalineResultAcrylic,
        [AlkalineWoolScale] = od.AlkalineWoolScale,
        [AlkalineResultWool] = od.AlkalineResultWool,

        [AcidChangeScale] = od.AcidChangeScale,
        [AcidResultChange] = od.AcidResultChange,
        [AcidAcetateScale] = od.AcidAcetateScale,
        [AcidResultAcetate] = od.AcidResultAcetate,
        [AcidCottonScale] = od.AcidCottonScale,
        [AcidResultCotton] = od.AcidResultCotton,
        [AcidNylonScale] = od.AcidNylonScale,
        [AcidResultNylon] = od.AcidResultNylon,
        [AcidPolyesterScale] = od.AcidPolyesterScale,
        [AcidResultPolyester] = od.AcidResultPolyester,
        [AcidAcrylicScale] = od.AcidAcrylicScale,
        [AcidResultAcrylic] = od.AcidResultAcrylic,
        [AcidWoolScale] = od.AcidWoolScale,
        [AcidResultWool] = od.AcidResultWool,

        [Remark] = od.Remark,
        [LastUpdate] = Concat(od.EditName, '-', pass1EditName.Name, ' ', pass1EditName.Extno),
        [Temperature] = cast(o.Temperature as varchar),
        [Time] = cast(o.Time as varchar),
        [Supplier] = ps.SuppID+'-'+s.AbbEN
from PerspirationFastness_Detail od with (nolock)
inner join PerspirationFastness o with (nolock) on o.ID = od.ID
left join PO_Supp_Detail psd with (nolock) on o.POID = psd.ID and od.SEQ1 = psd.SEQ1 and od.SEQ2 = psd.SEQ2
left join PO_Supp ps WITH (NOLOCK) on psd.ID = ps.ID and psd.Seq1 = ps.Seq1
left join supp s with (nolock) on ps.SuppID = s.ID
left join pass1 pass1EditName WITH(NOLOCK) on od.EditName = pass1EditName.ID
where   o.POID = @POID and o.TestNo = @TestNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, listPar);
        }

        public DataTable GetPerspirationFastness(string poID, string TestNo)
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
from    PerspirationFastness ov with (nolock)
left join [ExtendServer].PMSFile.dbo.PerspirationFastness oi with (nolock) on oi.ID=ov.ID
where   ov.POID = @poID and ov.TestNo = @TestNo
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
delete  [ExtendServer].PMSFile.dbo.PerspirationFastness where ID = (select ID from PerspirationFastness where POID = @poID and TestNo = @TestNo)
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
        ,SCIRefno_Color = po3.SCIRefno + ' ' +po3.ColorID
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
        ,pmsFile.TestBeforePicture
        ,pmsFile.TestAfterPicture
from PerspirationFastness_Detail cd WITH(NOLOCK)
left join PerspirationFastness c WITH(NOLOCK) on c.ID =  cd.ID
left join ExtendServer.PMSFile.dbo.PerspirationFastness pmsFile WITH(NOLOCK) on pmsFile.ID =  cd.ID
left join Orders o WITH(NOLOCK) on o.ID=c.POID
left join PO_Supp_Detail po3 WITH(NOLOCK) on c.POID = po3.ID 
	and cd.SEQ1 = po3.SEQ1 and cd.SEQ2 = po3.SEQ2
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
