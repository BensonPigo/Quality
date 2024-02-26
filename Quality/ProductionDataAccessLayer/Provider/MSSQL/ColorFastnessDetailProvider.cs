using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using MICS.DataAccessLayer.Interface;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using Sci;
using System.Data.SqlClient;
using System.Linq;
using ADOHelper.DBToolKit; 

namespace MICS.DataAccessLayer.Provider.MSSQL
{
    public class ColorFastnessDetailProvider : SQLDAL, IColorFastnessDetailProvider
    {
        #region 底層連線
        public ColorFastnessDetailProvider(string ConString) : base(ConString) { }
        public ColorFastnessDetailProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        public Fabric_ColorFastness_Detail_ViewModel Get_DetailBody(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SqlParameter para = new SqlParameter()
            {
                ParameterName = "@ID",
                SqlDbType = SqlDbType.VarChar,
                Value = ID
            };

            objParameter.Add(para);

            string sqlcmd = @"
select c.* 
    ,[Name] = p.Name
    ,[InspectionName] = Concat (Inspector, ' ', p.Name)
    ,o.BrandID
    ,TestBeforePicture = (select top 1 TestBeforePicture from SciPMSFile_ColorFastness pmsFile WITH(NOLOCK) where pmsFile.ID = c.ID and pmsFile.POID = c.POID and pmsFile.TestNo = c.TestNo)
    ,TestAfterPicture = (select top 1 TestAfterPicture from SciPMSFile_ColorFastness pmsFile WITH(NOLOCK) where pmsFile.ID = c.ID and pmsFile.POID = c.POID and pmsFile.TestNo = c.TestNo)
from ColorFastness c WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) on c.POID = o.ID  
left join pass1 p on c.Inspector = p.ID
where c.id = @ID
";
            var main = ExecuteList<ColorFastness_Result>(CommandType.Text, sqlcmd, objParameter);

            string sqlcmd2 = @"
select cd.ID
    , SubmitDate
    ,cd.ColorFastnessGroup
    ,Seq = CONCAT(cd.SEQ1,'-',cd.SEQ2)
    ,cd.Roll
    ,cd.Dyelot
    ,psd.Refno
    ,psd.SCIRefno
    ,ColorID = pc.SpecValue
    ,cd.Result  --所有檢驗過的匯總
    ,cd.changeScale
    ,cd.ResultChange
    ,cd.StainingScale
    ,cd.ResultStain
    ,cd.AcetateScale
    ,cd.ResultAcetate
    ,cd.CottonScale
    ,cd.ResultCotton
    ,cd.NylonScale
    ,cd.ResultNylon
    ,cd.PolyesterScale
    ,cd.ResultPolyester
    ,cd.AcrylicScale
    ,cd.ResultAcrylic
    ,cd.WoolScale
    ,cd.ResultWool
    ,cd.Remark
    ,[LastUpdate] = case 
        when cd.EditName !='' then CONCAT(cd.EditName,'-',pEdit.Name,pEdit.ExtNo)
        when cd.AddName !='' then CONCAT(cd.AddName,'-',pAdd.Name,pAdd.ExtNo)
        else '' end
    ,cd.AddDate,cd.AddName,cd.EditDate,cd.EditName
from ColorFastness_Detail cd WITH(NOLOCK)
left join ColorFastness c WITH(NOLOCK) on c.ID =  cd.ID
left join PO_Supp_Detail psd WITH(NOLOCK) on c.POID = psd.ID and cd.SEQ1 = psd.SEQ1 and cd.SEQ2 = psd.SEQ2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Pass1 pEdit WITH(NOLOCK) on pEdit.ID = cd.EditName
left join pass1 pAdd WITH(NOLOCK) on pAdd.ID = cd.AddName
where cd.ID = @ID
order by cd.id,cd.ColorFastnessGroup
";
            var detail = ExecuteList<Fabric_ColorFastness_Detail_Result>(CommandType.Text, sqlcmd2, objParameter);

            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel
            {
                Main = main.FirstOrDefault(),
                Details = detail.ToList(),
            };

            return result;
        }

        public IList<ColorFastness_Excel> GetExcel(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            string sqlcmd = @"
select [ReportNo] = c.ID
		,cd.SubmitDate
        ,o.SeasonID
        ,o.BrandID
        ,o.StyleID
        ,c.POID
		,c.Article
		,c.Temperature
		,c.Cycle
		,c.CycleTime
		,c.Detergent
		,c.Machine
		,c.Drying
        ,ColorFastnessResult = c.Result
		,cd.SEQ1
		,cd.SEQ2
		,cd.Roll
        ,cd.Dyelot
        ,SCIRefno_Color = psd.SCIRefno + ' ' + pc.SpecValue       
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
        ,cd.Result
		,cd.Remark
		,cd.StainingScale	
        ,cd.ResultStain          
        ,Signature = (select t.Signature from Technician t where t.ID = c.Inspector and t.Junk = 0)
        ,TestBeforePicture = (select top 1 TestBeforePicture from SciPMSFile_ColorFastness pmsFile WITH(NOLOCK) where pmsFile.ID = c.ID and pmsFile.POID = c.POID and pmsFile.TestNo = c.TestNo)
        ,TestAfterPicture = (select top 1 TestAfterPicture from SciPMSFile_ColorFastness pmsFile WITH(NOLOCK) where pmsFile.ID = c.ID and pmsFile.POID = c.POID and pmsFile.TestNo = c.TestNo)
		,[Checkby] = isnull(pEdit.Name, pAdd.Name)
from ColorFastness_Detail cd WITH(NOLOCK)
left join ColorFastness c WITH(NOLOCK) on c.ID =  cd.ID
left join Orders o WITH(NOLOCK) on o.ID=c.POID
left join PO_Supp_Detail psd WITH(NOLOCK) on c.POID = psd.ID and cd.SEQ1 = psd.SEQ1 and cd.SEQ2 = psd.SEQ2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Pass1 pEdit WITH(NOLOCK) on pEdit.ID = cd.EditName
left join pass1 pAdd WITH(NOLOCK) on pAdd.ID = cd.AddName
where cd.ID = @ID
order by cd.SubmitDate
";
            var detail = ExecuteList<ColorFastness_Excel>(CommandType.Text, sqlcmd, objParameter);

            return detail.Any() ? detail : new List<ColorFastness_Excel>();
        }
        public IList<PO_Supp_Detail> Get_Seq(string POID, string Seq1,string Seq2)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@POID", DbType.String, POID } ,
                { "@Seq1", DbType.String, Seq1 } ,
                { "@Seq2", DbType.String, Seq2 } ,
            };

            string sqlcmd = @"
select POID = psd.ID
,SEQ = CONCAT(psd.SEQ1,'-',psd.SEQ2) 
,psd.Seq1
,psd.Seq2
,psd.SCIRefno
,psd.Refno
,ColorID = pc.SpecValue
from PO_Supp_Detail psd WITH(NOLOCK)
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
where psd.ID = @POID
and psd.FabricType = 'F'
";
            if (!string.IsNullOrEmpty(Seq1))
            {
                sqlcmd += " and psd.Seq1 = @Seq1" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(Seq2))
            {
                sqlcmd += " and psd.Seq2 = @Seq2" + Environment.NewLine;
            }

            return ExecuteList<PO_Supp_Detail>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<FtyInventory> Get_Roll(string POID, string Seq1, string Seq2)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@POID", DbType.String, POID } ,
                { "@Seq1", DbType.String, Seq1 } ,
                { "@Seq2", DbType.String, Seq2 } ,
            };

            string sqlcmd = @"
select POID,Seq1,Seq2
,Roll,Dyelot
from FtyInventory WITH(NOLOCK)
where POID = @POID
and Seq1 = @Seq1
and Seq2 = @Seq2
";
            if (!string.IsNullOrEmpty(Seq1))
            {
                sqlcmd += " and Seq1 = @Seq1" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(Seq2))
            {
                sqlcmd += " and Seq2 = @Seq2" + Environment.NewLine;
            }

            return ExecuteList<FtyInventory>(CommandType.Text, sqlcmd, objParameter);
        }

        public DataTable Get_SubmitDate(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            string sqlcmd = @"
select distinct CONVERT(varchar(100), SubmitDate, 111) as SubmitDate 
from ColorFastness_Detail WITH (NOLOCK) 
where id = @ID
";
            return ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);
        }

        public bool Save_ColorFastness(Fabric_ColorFastness_Detail_ViewModel sources, string Mdivision,string UserID, out string NewID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@POID", sources.Main.POID } ,
                { "@InspDate", sources.Main.InspDate } ,
                { "@Article", sources.Main.Article } ,
                { "@Result", sources.Main.Result } ,
                { "@Status", sources.Main.Status } ,
                { "@Inspector", sources.Main.Inspector ?? ""} ,
                { "@Remark", sources.Main.Remark ?? ""} ,
                { "@Temperature", sources.Main.Temperature } ,
                { "@Cycle", sources.Main.Cycle } ,
                { "@CycleTime", sources.Main.CycleTime } ,
                { "@Detergent", sources.Main.Detergent ?? ""} ,
                { "@Machine", sources.Main.Machine} ,
                { "@Drying", sources.Main.Drying ?? ""} ,
                { "@UserID", UserID } ,
            };

            if (sources.Main.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", sources.Main.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }

            if (sources.Main.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", sources.Main.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            string sqlcmd = string.Empty;
            NewID = string.Empty;
            string ID = string.Empty;
            #region save Main


            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, $@"select Max(testno) as testMaxNo from ColorFastness WITH (NOLOCK) where poid='{sources.Main.POID}'", new SQLParameterCollection());
            int testMaxNo = MyUtility.Convert.GetInt(dt.Rows[0]["testMaxNo"]);
            objParameter.Add("@TestNo", testMaxNo + 1);

            if (sources.Main.ID != null && !string.IsNullOrEmpty(sources.Main.ID))
            {
                ID = sources.Main.ID;
                objParameter.Add(new SqlParameter($"@ID", sources.Main.ID));
                objParameter.Add(new SqlParameter($"@OriTestNo", sources.Main.TestNo));
                // update 
                sqlcmd += @"
SET XACT_ABORT ON

update ColorFastness
set	  [InspDate] = @InspDate
      ,[Article] = @Article
      ,[Status] = @Status
      ,[Inspector] = @Inspector
      ,[Remark] = @Remark
      ,[EditName] = @UserID
      ,[EditDate] = GetDate()
      ,[Temperature] = @Temperature
      ,[Cycle] = @Cycle
      ,[CycleTime] = @CycleTime
      ,[Detergent] = @Detergent
      ,[Machine] = @Machine
      ,[Drying] = @Drying
      -----2022/01/10 PMSFile上線，因此去掉Image寫入DB的部分
where ID = @ID
and POID = @POID 
and TestNo = @OriTestNo

if not exists (select 1 from SciPMSFile_ColorFastness where ID = @ID and POID = @POID and TestNo = @OriTestNo)
begin
    INSERT INTO SciPMSFile_ColorFastness (ID,POID,TestNo,TestBeforePicture,TestAfterPicture)
    VALUES (@ID,@POID,@OriTestNo,@TestBeforePicture,@TestAfterPicture)
end
else
begin
    update SciPMSFile_ColorFastness
    set	   [TestBeforePicture] = @TestBeforePicture
          ,[TestAfterPicture] = @TestAfterPicture
    where ID = @ID
    and POID = @POID 
    and TestNo = @OriTestNo
end

exec UpdateInspPercent 'LabColorFastness', @POID
" + Environment.NewLine;
            }
            else
            {
                NewID = GetID(Mdivision + "CF", "ColorFastness", DateTime.Today, 2, "ID");
                ID = NewID;
                objParameter.Add(new SqlParameter($"@ID", NewID));
                sqlcmd += @"
SET XACT_ABORT ON

----2022/01/10 PMSFile上線，因此去掉Image寫入原本DB的部分
insert into ColorFastness(ID,POID,TestNo,InspDate,Article,Status,Inspector,Remark,addName,addDate,Temperature,Cycle,CycleTime,Detergent,Machine,Drying)
values(@ID ,@POID,@TestNo,GETDATE(),@Article,'New',@UserID,@Remark,@UserID,GETDATE(),@Temperature,@Cycle,@CycleTime,@Detergent,@Machine,@Drying)

insert into SciPMSFile_ColorFastness(ID,POID,TestNo,TestBeforePicture,TestAfterPicture)
values(@ID,@POID,@TestNo,@TestBeforePicture,@TestAfterPicture)


";
            }
            #endregion

            Fabric_ColorFastness_Detail_ViewModel oldData = Get_DetailBody(ID);
            List<Fabric_ColorFastness_Detail_Result> oldDetailData = oldData.Details;

            List<Fabric_ColorFastness_Detail_Result> needUpdateDetailList =
                ToolKit.PublicClass.CompareListValue<Fabric_ColorFastness_Detail_Result>(
                    sources.Details,
                    oldDetailData,
                    "ID,ColorFastnessGroup,SEQ1,SEQ2",
                    "Roll,Dyelot,Remark,SubmitDate,Result,changeScale,ResultChange,StainingScale,ResultStain,AcetateScale,ResultAcetate,CottonScale,ResultCotton,NylonScale,ResultNylon,PolyesterScale,ResultPolyester,AcrylicScale,ResultAcrylic,WoolScale,ResultWool");

            #region save Details

            string insertDetail = $@"
insert into ColorFastness_Detail 
(        [ID]
        ,[ColorFastnessGroup]
        ,[SEQ1]
        ,[SEQ2]
        ,[Roll]
        ,[Dyelot]
        ,[Result]
        ,[changeScale]
        ,[ResultChange]
        ,[Remark]
        ,[AddName]
        ,[AddDate]
        ,[SubmitDate]
        ,[StainingScale]
        ,[ResultStain]
        ,[AcetateScale]
        ,[ResultAcetate]
        ,[CottonScale]
        ,[ResultCotton]
        ,[NylonScale]
        ,[ResultNylon]
        ,[PolyesterScale]
        ,[ResultPolyester]
        ,[AcrylicScale]
        ,[ResultAcrylic]
        ,[WoolScale]
        ,[ResultWool]
) 
values
(
         @ID
        ,@ColorFastnessGroup
        ,@Seq1
        ,@Seq2
        ,@Roll
        ,@Dyelot
        ,@Result
        ,@changeScale
        ,@ResultChange
        ,@Remark
        ,@UserID
        ,GetDate()      
        ,@SubmitDate
        ,@StainingScale
        ,@ResultStain
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
        ,@ResultWool
)

declare @POID varchar(13) = (select POID from ColorFastness WITH(NOLOCK) where ID = @ID)
exec UpdateInspPercent 'LabColorFastness', @POID
";
            string deleteDetail = $@"
delete from ColorFastness_Detail 
where id = @ID
and ColorFastnessGroup = @ColorFastnessGroup
and SEQ1 = @SEQ1
and SEQ2 = @SEQ2

declare @POID varchar(13) = (select POID from ColorFastness WITH(NOLOCK) where ID = @ID)
exec UpdateInspPercent 'LabColorFastness', @POID
";
            string updateDetail = $@"
update ColorFastness_Detail
set 
       [Roll] = @Roll
        ,[Dyelot] = @Dyelot
        ,[Result] = @Result
        ,[changeScale] = @changeScale
        ,[ResultChange] = @ResultChange
        ,[Remark] = @Remark
        ,[EditName] = @UserID
        ,[EditDate] = GetDate()
        ,[SubmitDate] = @SubmitDate
        ,StainingScale=@StainingScale
        ,ResultStain=@ResultStain
        ,AcetateScale=@AcetateScale
        ,ResultAcetate=@ResultAcetate
        ,CottonScale=@CottonScale
        ,ResultCotton=@ResultCotton
        ,NylonScale=@NylonScale
        ,ResultNylon=@ResultNylon
        ,PolyesterScale=@PolyesterScale
        ,ResultPolyester=@ResultPolyester
        ,AcrylicScale=@AcrylicScale
        ,ResultAcrylic=@ResultAcrylic
        ,WoolScale=@WoolScale
        ,ResultWool=@ResultWool
where ID = @ID
and ColorFastnessGroup = @ColorFastnessGroup
and SEQ1 = @Seq1
and SEQ2 = @Seq2

declare @POID varchar(13) = (select POID from ColorFastness WITH(NOLOCK) where ID = @ID)
exec UpdateInspPercent 'LabColorFastness', @POID
";

            ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
            foreach (var detailItem in needUpdateDetailList)
            {   
                SQLParameterCollection listDetailPar = new SQLParameterCollection();
                string DetailResult = (detailItem.ResultChange.EqualString("Fail")
                    || detailItem.ResultStain.EqualString("Fail")
                    || detailItem.ResultAcetate.EqualString("Fail") 
                    || detailItem.ResultCotton.EqualString("Fail") 
                    || detailItem.ResultNylon.EqualString("Fail") 
                    || detailItem.ResultPolyester.EqualString("Fail") 
                    || detailItem.ResultAcrylic.EqualString("Fail") 
                    || detailItem.ResultWool.EqualString("Fail")) ? "Fail" : "Pass";

                switch (detailItem.StateType)
                {
                    case DatabaseObject.Public.CompareStateType.Add:
                        listDetailPar.Add(new SqlParameter($"@ID", ID));
                        listDetailPar.Add(new SqlParameter($"@ColorFastnessGroup", string.IsNullOrEmpty(detailItem.ColorFastnessGroup) ? "" : detailItem.ColorFastnessGroup));
                        listDetailPar.Add(new SqlParameter($"@Seq1", detailItem.SEQ1));
                        listDetailPar.Add(new SqlParameter($"@Seq2", detailItem.SEQ2));
                        listDetailPar.Add(new SqlParameter($"@Roll", string.IsNullOrEmpty(detailItem.Roll) ? "" : detailItem.Roll));
                        listDetailPar.Add(new SqlParameter($"@Dyelot", string.IsNullOrEmpty(detailItem.Dyelot) ? "" : detailItem.Dyelot));
                        listDetailPar.Add(new SqlParameter($"@Result", DetailResult));
                        listDetailPar.Add(new SqlParameter($"@changeScale", detailItem.changeScale));
                        listDetailPar.Add(new SqlParameter($"@ResultChange", detailItem.ResultChange));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? ""));
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                        listDetailPar.Add($"@SubmitDate", DbType.Date, detailItem.SubmitDate);
                        listDetailPar.Add(new SqlParameter($"@StainingScale", detailItem.StainingScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultStain", detailItem.ResultStain ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcetateScale", detailItem.AcetateScale?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultAcetate", detailItem.ResultAcetate ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@CottonScale", detailItem.CottonScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultCotton", detailItem.ResultCotton ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@NylonScale", detailItem.NylonScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultNylon", detailItem.ResultNylon ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@PolyesterScale", detailItem.PolyesterScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultPolyester", detailItem.ResultPolyester ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcrylicScale", detailItem.AcrylicScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultAcrylic", detailItem.ResultAcrylic ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WoolScale", detailItem.WoolScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultWool", detailItem.ResultWool ?? string.Empty));

                        ExecuteNonQuery(CommandType.Text, insertDetail, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Edit:
                        listDetailPar.Add(new SqlParameter($"@ID", ID));
                        listDetailPar.Add(new SqlParameter($"@ColorFastnessGroup", string.IsNullOrEmpty(detailItem.ColorFastnessGroup) ? "" : detailItem.ColorFastnessGroup));
                        listDetailPar.Add(new SqlParameter($"@Seq1", detailItem.SEQ1));
                        listDetailPar.Add(new SqlParameter($"@Seq2", detailItem.SEQ2));
                        listDetailPar.Add(new SqlParameter($"@Roll", string.IsNullOrEmpty(detailItem.Roll) ? "" : detailItem.Roll));
                        listDetailPar.Add(new SqlParameter($"@Dyelot", string.IsNullOrEmpty(detailItem.Dyelot) ? "" : detailItem.Dyelot));
                        listDetailPar.Add(new SqlParameter($"@Result", DetailResult));
                        listDetailPar.Add(new SqlParameter($"@changeScale", detailItem.changeScale));
                        listDetailPar.Add(new SqlParameter($"@ResultChange", detailItem.ResultChange));
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark ?? ""));
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                        listDetailPar.Add($"@SubmitDate", DbType.Date, detailItem.SubmitDate);
                        listDetailPar.Add(new SqlParameter($"@StainingScale", detailItem.StainingScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultStain", detailItem.ResultStain ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcetateScale", detailItem.AcetateScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultAcetate", detailItem.ResultAcetate ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@CottonScale", detailItem.CottonScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultCotton", detailItem.ResultCotton ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@NylonScale", detailItem.NylonScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultNylon", detailItem.ResultNylon ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@PolyesterScale", detailItem.PolyesterScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultPolyester", detailItem.ResultPolyester ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@AcrylicScale", detailItem.AcrylicScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultAcrylic", detailItem.ResultAcrylic ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@WoolScale", detailItem.WoolScale ?? string.Empty));
                        listDetailPar.Add(new SqlParameter($"@ResultWool", detailItem.ResultWool ?? string.Empty));
                        ExecuteNonQuery(CommandType.Text, updateDetail, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Delete:
                        listDetailPar.Add(new SqlParameter($"@ID", ID));
                        listDetailPar.Add(new SqlParameter($"@ColorFastnessGroup", string.IsNullOrEmpty(detailItem.ColorFastnessGroup) ? "" : detailItem.ColorFastnessGroup));
                        listDetailPar.Add(new SqlParameter($"@Seq1", detailItem.SEQ1));
                        listDetailPar.Add(new SqlParameter($"@Seq2", detailItem.SEQ2));

                        ExecuteNonQuery(CommandType.Text, deleteDetail, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.None:
                        break;
                    default:
                        break;
                }
            }

            #endregion

            return true;
        }

        public bool Delete_ColorFastness_Detail(string ID, List<Fabric_ColorFastness_Detail_Result> source)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            IList<Fabric_ColorFastness_Detail_Result> dbDetail = Get_DetailBody(ID).Details;
            string sqlcmd = string.Empty;

            int idx = 1;
            foreach (var item in dbDetail)
            {
                objParameter.Add(new SqlParameter($"@ID{idx}", item.ID));
                objParameter.Add(new SqlParameter($"@ColorFastnessGroup{idx}", string.IsNullOrEmpty(item.ColorFastnessGroup) ? "" : item.ColorFastnessGroup));
                objParameter.Add(new SqlParameter($"@SEQ1{idx}", item.SEQ1));
                objParameter.Add(new SqlParameter($"@SEQ2{idx}", item.SEQ2));

                if (!source.Where(x => x.ColorFastnessGroup.EqualString(item.ColorFastnessGroup.ToString())
                    && x.SEQ1.EqualString(item.SEQ1.ToString())
                    && x.SEQ2.EqualString(item.SEQ2.ToString())
                ).Any())
                {   
                    sqlcmd += $@"
delete from ColorFastness_Detail 
where id = @ID{idx} 
and ColorFastnessGroup = @ColorFastnessGroup{idx}
and SEQ1 = @SEQ1{idx} 
and SEQ2 = @SEQ2{idx} 

declare @POID varchar(13) = (select POID from ColorFastness WITH(NOLOCK) where ID = @ID)
exec UpdateInspPercent 'LabColorFastness', @POID
" + Environment.NewLine;
                    idx++;
                }
            }

            if (string.IsNullOrEmpty(sqlcmd))
            {
                return true;
            }
            else
            {
                return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
            }
        }


        #endregion

        public class PO_Supp_Detail
        {
            public string POID { get; set; }
            public string SEQ { get; set; }
            public string Seq1 { get; set; }
            public string Seq2 { get; set; }
            public string SCIRefno { get; set; }
            public string Refno { get; set; }
            public string ColorID { get; set; }
        }

        public class FtyInventory
        {
            public string POID { get; set; }
            public string Seq1 { get; set; }
            public string Seq2 { get; set; }
            public string Roll { get; set; }
            public string Dyelot { get; set; }
        }
    }
}
