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
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            string sqlcmd = @"
select c.* 
,[Name] = p.Name
,[InspectionName] = Concat (Inspector, ' ', p.Name)
from ColorFastness c WITH(NOLOCK)
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
,po3.Refno,po3.SCIRefno,po3.ColorID

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

,cd.ResultStain,cd.Remark
,[LastUpdate] = case 
    when cd.EditName !='' then CONCAT(cd.EditName,'-',pEdit.Name,pEdit.ExtNo)
    when cd.AddName !='' then CONCAT(cd.AddName,'-',pAdd.Name,pAdd.ExtNo)
    else '' end
,cd.AddDate,cd.AddName,cd.EditDate,cd.EditName
from ColorFastness_Detail cd WITH(NOLOCK)
left join ColorFastness c WITH(NOLOCK) on c.ID =  cd.ID
left join PO_Supp_Detail po3 WITH(NOLOCK) on c.POID = po3.ID 
	and cd.SEQ1 = po3.SEQ1 and cd.SEQ2 = po3.SEQ2
left join Pass1 pEdit WITH(NOLOCK) on pEdit.ID = cd.EditName
left join pass1 pAdd WITH(NOLOCK) on pAdd.ID = cd.AddName
where cd.ID = @ID
order by cd.id,cd.ColorFastnessGroup
";
            var detail = ExecuteList<Fabric_ColorFastness_Detail_Result>(CommandType.Text, sqlcmd2, objParameter);

            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel
            {
                Main = main.FirstOrDefault(),
                Detail = detail.ToList(),
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
select cd.SubmitDate
        ,o.SeasonID
        ,o.BrandID
        ,o.StyleID
        ,c.POID
        ,cd.Roll
        ,cd.Dyelot
        ,SCIRefno_Color = po3.SCIRefno + ' ' +po3.ColorID
        ,c.Temperature
        ,c.CycleTime
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
        ,pmsFile.TestBeforePicture
        ,pmsFile.TestAfterPicture
from ColorFastness_Detail cd WITH(NOLOCK)
left join ColorFastness c WITH(NOLOCK) on c.ID =  cd.ID
left join ExtendServer.PMSFile.dbo.ColorFastness pmsFile WITH(NOLOCK) on pmsFile.ID =  cd.ID
left join Orders o WITH(NOLOCK) on o.ID=c.POID
left join PO_Supp_Detail po3 WITH(NOLOCK) on c.POID = po3.ID 
	and cd.SEQ1 = po3.SEQ1 and cd.SEQ2 = po3.SEQ2
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
select POID = ID
,SEQ = CONCAT(SEQ1,'-',SEQ2) 
,Seq1,Seq2
,SCIRefno
,Refno
,ColorID
from PO_Supp_Detail WITH(NOLOCK)
where ID = @POID
and FabricType = 'F'
";
            if (!string.IsNullOrEmpty(Seq1))
            {
                sqlcmd += " and Seq1 = @Seq1" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(Seq2))
            {
                sqlcmd += " and Seq2 = @Seq2" + Environment.NewLine;
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
                { "@Inspector", sources.Main.Inspector } ,
                { "@Remark", sources.Main.Remark } ,
                { "@Temperature", sources.Main.Temperature } ,
                { "@Cycle", sources.Main.Cycle } ,
                { "@CycleTime", sources.Main.CycleTime } ,
                { "@Detergent", sources.Main.Detergent } ,
                { "@Machine", sources.Main.Machine } ,
                { "@Drying", sources.Main.Drying } ,
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
                // update 
                sqlcmd += @"
SET XACT_ABORT ON

update ColorFastness
set	   [POID] = @POID
      ,[InspDate] = @InspDate
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
      ,[TestBeforePicture] = @TestBeforePicture
      ,[TestAfterPicture] = @TestAfterPicture
where ID = @ID

update [ExtendServer].PMSFile.dbo.ColorFastness
set	   [TestBeforePicture] = @TestBeforePicture
      ,[TestAfterPicture] = @TestAfterPicture
where ID = @ID


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

insert into ColorFastness(ID,POID,TestNo,InspDate,Article,Status,Inspector,Remark,addName,addDate,Temperature,Cycle,CycleTime,Detergent,Machine,Drying,TestBeforePicture,TestAfterPicture)
values(@ID ,@POID,@TestNo,GETDATE(),@Article,'New',@UserID,@Remark,@UserID,GETDATE(),@Temperature,@Cycle,@CycleTime,@Detergent,@Machine,@Drying,@TestBeforePicture,@TestAfterPicture)

insert into [ExtendServer].PMSFile.dbo.ColorFastness(ID,TestBeforePicture,TestAfterPicture)
values(@ID ,@TestBeforePicture,@TestAfterPicture)


";
            }
            #endregion

            Fabric_ColorFastness_Detail_ViewModel oldData = Get_DetailBody(ID);
            List<Fabric_ColorFastness_Detail_Result> oldDetailData = oldData.Detail;

            List<Fabric_ColorFastness_Detail_Result> needUpdateDetailList =
                ToolKit.PublicClass.CompareListValue<Fabric_ColorFastness_Detail_Result>(
                    sources.Detail,
                    oldDetailData,
                    "ID,ColorFastnessGroup,SEQ1,SEQ2",
                    "Roll,Dyelot,changeScale,StainingScale,Remark,SubmitDate,ResultChange,ResultStain");

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
)";
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
";

            ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
            foreach (var detailItem in needUpdateDetailList)
            {   
                SQLParameterCollection listDetailPar = new SQLParameterCollection();
                string DetailResult = (detailItem.ResultChange.EqualString("Pass") && detailItem.ResultStain.EqualString("Pass")) ? "Pass" : "Fail";

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
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark));
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                        listDetailPar.Add($"@SubmitDate", DbType.Date, detailItem.SubmitDate);

                        listDetailPar.Add(new SqlParameter($"@AcetateScale", detailItem.AcetateScale));
                        listDetailPar.Add(new SqlParameter($"@ResultAcetate", detailItem.ResultAcetate));
                        listDetailPar.Add(new SqlParameter($"@CottonScale", detailItem.CottonScale));
                        listDetailPar.Add(new SqlParameter($"@ResultCotton", detailItem.ResultCotton));
                        listDetailPar.Add(new SqlParameter($"@NylonScale", detailItem.NylonScale));
                        listDetailPar.Add(new SqlParameter($"@ResultNylon", detailItem.ResultNylon));
                        listDetailPar.Add(new SqlParameter($"@PolyesterScale", detailItem.PolyesterScale));
                        listDetailPar.Add(new SqlParameter($"@ResultPolyester", detailItem.ResultPolyester));
                        listDetailPar.Add(new SqlParameter($"@AcrylicScale", detailItem.AcrylicScale));
                        listDetailPar.Add(new SqlParameter($"@ResultAcrylic", detailItem.ResultAcrylic));
                        listDetailPar.Add(new SqlParameter($"@WoolScale", detailItem.WoolScale));
                        listDetailPar.Add(new SqlParameter($"@ResultWool", detailItem.ResultWool));

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
                        listDetailPar.Add(new SqlParameter($"@Remark", detailItem.Remark));
                        listDetailPar.Add(new SqlParameter($"@UserID", UserID));
                        listDetailPar.Add($"@SubmitDate", DbType.Date, detailItem.SubmitDate);

                        listDetailPar.Add(new SqlParameter($"@AcetateScale", detailItem.AcetateScale));
                        listDetailPar.Add(new SqlParameter($"@ResultAcetate", detailItem.ResultAcetate));
                        listDetailPar.Add(new SqlParameter($"@CottonScale", detailItem.CottonScale));
                        listDetailPar.Add(new SqlParameter($"@ResultCotton", detailItem.ResultCotton));
                        listDetailPar.Add(new SqlParameter($"@NylonScale", detailItem.NylonScale));
                        listDetailPar.Add(new SqlParameter($"@ResultNylon", detailItem.ResultNylon));
                        listDetailPar.Add(new SqlParameter($"@PolyesterScale", detailItem.PolyesterScale));
                        listDetailPar.Add(new SqlParameter($"@ResultPolyester", detailItem.ResultPolyester));
                        listDetailPar.Add(new SqlParameter($"@AcrylicScale", detailItem.AcrylicScale));
                        listDetailPar.Add(new SqlParameter($"@ResultAcrylic", detailItem.ResultAcrylic));
                        listDetailPar.Add(new SqlParameter($"@WoolScale", detailItem.WoolScale));
                        listDetailPar.Add(new SqlParameter($"@ResultWool", detailItem.ResultWool));
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
            IList<Fabric_ColorFastness_Detail_Result> dbDetail = Get_DetailBody(ID).Detail;
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

        public IList<ColorFastness_Detail> Get(ColorFastness_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,ColorFastnessGroup"+ Environment.NewLine);
            SbSql.Append("        ,SEQ1"+ Environment.NewLine);
            SbSql.Append("        ,SEQ2"+ Environment.NewLine);
            SbSql.Append("        ,Roll"+ Environment.NewLine);
            SbSql.Append("        ,Dyelot"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,changeScale"+ Environment.NewLine);
            SbSql.Append("        ,StainingScale"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,SubmitDate"+ Environment.NewLine);
            SbSql.Append("        ,ResultChange"+ Environment.NewLine);
            SbSql.Append("        ,ResultStain"+ Environment.NewLine);
            SbSql.Append("FROM [ColorFastness_Detail]"+ Environment.NewLine);



            return ExecuteList<ColorFastness_Detail>(CommandType.Text, SbSql.ToString(), objParameter);
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
