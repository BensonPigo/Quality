using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel;
using DatabaseObject.RequestModel;
using System.Linq;
using System.Data.SqlClient;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentTestProvider : SQLDAL, IGarmentTestProvider
    {
        #region 底層連線
        public GarmentTestProvider(string ConString) : base(ConString) { }
        public GarmentTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public IList<Style> GetStyleID()
        {
            string sqlcmd = @"select distinct ID from Style where Junk = 0";

            return ExecuteList<Style>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }

        public IList<Brand> GetBrandID()
        {
            string sqlcmd = @"select distinct ID from Brand where Junk = 0";

            return ExecuteList<Brand>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }

        public IList<Season> GetSeasonID()
        {
            string sqlcmd = @"select distinct ID from Season where Junk = 0";

            return ExecuteList<Season>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }

        public IList<GarmentTest> GetArticle(GarmentTest_ViewModel filter)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, filter.BrandID } ,
                { "@StyleID", DbType.String, filter.StyleID } ,
                { "@SeasonID", DbType.String, filter.SeasonID} ,
            };
            string sqlcmd = @"
select distinct Article from GarmentTest
where 1=1
and BrandID = @BrandID
and StyleID = @StyleID
and SeasonID = @SeasonID";

            return ExecuteList<GarmentTest>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<GarmentTest_ViewModel> Get_GarmentTest(GarmentTest_Request filter)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, filter.Brand } ,
                { "@StyleID", DbType.String, filter.Style } ,
                { "@SeasonID", DbType.String, filter.Season } ,
                { "@Article", DbType.String, filter.Article } ,
                { "@ID", DbType.String, filter.ID } ,
            };
            string sqlcmd = @"
select g.ID
,g.StyleID
,g.BrandID
,g.Article
,g.SeasonID
,g.FirstOrderID
,g.DeadLine
,g.SewingInline
,g.OrderID
,g.MDivisionid
,[MinSciDelivery] = GetSCI.MinSciDelivery
,[MinBuyerDelivery] = GetSCI.MinBuyerDelivery
,[SeamBreakageResult] = case when g.SeamBreakageResult = 'P' then 'Pass'
						     when g.SeamBreakageResult = 'F' then 'Fail' 
                             else '' end
,[OdourResult] = case when g.OdourResult = 'P' then 'Pass'
					  when g.OdourResult = 'F' then 'Fail' 
                      else '' end
,[WashResult] = case when g.WashResult = 'P' then 'Pass'
				     when g.WashResult = 'F' then 'Fail' 
                     else '' end
,[WashName] = IIF(WashName.Value is null,'701','710')
,[SpecialMark] = SpecialMark.Value
,g.Result, g.Date,g.Remark
,[GarmentTestAddName] = CONCAT(g.AddName,'-',CreatBy.Name,'',g.AddDate)
,[GarmentTestEditName] = CONCAT(g.EditName,'-',EditBy.Name,'',g.EditDate)
,g.AddName,g.EditName
from GarmentTest g
left join Pass1 CreatBy on CreatBy.ID = g.AddName
left join Pass1 EditBy on EditBy.ID = g.EditName
outer apply(
	select MinBuyerDelivery,MinSciDelivery
	from dbo.GetSCI(g.FirstOrderID,'')
) GetSCI
outer apply(
	select Value =  r.Name 
	from Style s
	inner join Reason r on s.SpecialMark = r.ID and r.ReasonTypeID = 'Style_SpecialMark'
	where s.ID = g.StyleID
	and s.BrandID = g.BrandID
	and s.SeasonID = g.SeasonID
)SpecialMark
outer apply(
	select Value =  r.Name 
	from Style s
	inner join Reason r on s.SpecialMark = r.ID and r.ReasonTypeID = 'Style_SpecialMark'
	where s.ID = g.StyleID
	and s.BrandID = g.BrandID
	and s.SeasonID = g.SeasonID
	and r.Name in ('MATCH TEAMWEAR','BASEBALL ON FIELD','SOFTBALL ON FIELD','TRAINING TEAMWEAR','LACROSSE ONFIELD','AMERIC. FOOT. ON-FIELD','TIRO','BASEBALL OFF FIELD','NCAA ON ICE','ON-COURT','BBALL PERFORMANCE','BRANDED BLANKS','SLD ON-FIELD','NHL ON ICE','SLD ON-COURT')
)WashName
where 1=1
";
            if (!string.IsNullOrEmpty(filter.MDivisionid))
            {
                objParameter.Add("@MDivisionid", DbType.String, filter.MDivisionid);
                sqlcmd += " and g.MDivisionid = @MDivisionid" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(filter.Brand))
            {   
                sqlcmd += " and g.BrandID = @BrandID" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(filter.Style))
            {
                sqlcmd += " and g.StyleID = @StyleID" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(filter.Season))
            {
                sqlcmd += " and g.SeasonID = @SeasonID" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(filter.Article))
            {
                sqlcmd += " and g.Article = @Article" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(filter.ID))
            {
                sqlcmd += " and g.ID = @ID" + Environment.NewLine;
            }

            return ExecuteList<GarmentTest_ViewModel>(CommandType.Text, sqlcmd, objParameter);
        }

        public bool Save_GarmentTest(GarmentTest_ViewModel master, List<GarmentTest_Detail> detail, string UserID)
        {
            bool result = true;

            foreach (var item in detail)
            {
                string sql_Shrinkage_Chk = $"select 1 from Production.dbo.GarmentTest_Detail_Shrinkage with(nolock) where id = '{master.ID}' and NO = '{item.No}'";
                DataTable dtChk_Shrinkage = ExecuteDataTableByServiceConn(CommandType.Text, sql_Shrinkage_Chk, new SQLParameterCollection());

                string sql_detail = $@"select 1 from GarmentTest_Detail with(nolock) where  id = '{master.ID}' and NO = '{item.No}'";
                DataTable dtDetail = ExecuteDataTableByServiceConn(CommandType.Text, sql_detail, new SQLParameterCollection());

                if (dtChk_Shrinkage!= null && dtChk_Shrinkage.Rows.Count == 0)
                {
                    #region insertShrinkage
                    SQLParameterCollection objParameter1 = new SQLParameterCollection
                    {
                        { "@ID", DbType.String, master.ID} ,
                        { "@BrandID", DbType.String, master.BrandID } ,
                        { "@No", DbType.String, item.No} ,
                    };

                    string insertShrinkage = $@"
select sl.Location
into #Location1
from GarmentTest gt with(nolock)
inner join style s with(nolock) on s.id = gt.StyleID
inner join Style_Location sl with(nolock) on sl.styleukey = s.ukey
where gt.id = @ID and sl.Location !='B'
group by sl.Location
order by sl.Location desc

CREATE TABLE #type1([type] [varchar](20),seq numeric(6,0))
insert into #type1 values('Chest Width',1)
insert into #type1 values('Sleeve Width',2)
insert into #type1 values('Sleeve Length',3)
insert into #type1 values('Back Length',4)
insert into #type1 values('Hem Opening',5)
---
select distinct sl.Location
into #Location2
from GarmentTest gt with(nolock)
inner join style s with(nolock) on s.id = gt.StyleID
inner join Style_Location sl with(nolock) on sl.styleukey = s.ukey
where gt.id = @ID and sl.Location ='B'

CREATE TABLE #type2([type] [varchar](20),seq numeric(6,0))
insert into #type2 values('Waistband (relax)',1)
insert into #type2 values('Hip Width',2)
insert into #type2 values('Thigh Width',3)
insert into #type2 values('Side Seam',4)
insert into #type2 values('Leg Opening',5)

select sl.Location,s.BrandID
into #Location_S
from GarmentTest gt with(nolock)
inner join style s with(nolock) on s.id = gt.StyleID
inner join Style_Location sl with(nolock) on sl.styleukey = s.ukey
where gt.id = @ID
group by sl.Location,s.BrandID
order by sl.Location desc

declare @location_combo varchar(15) = 
(select LocationList = Stuff((
	select concat(',',Location)
	from (
			select 	distinct
				Location
			from #Location_S d
		) s 
			order by Location asc
	for xml path ('')
) , 1, 1, ''))

if @location_combo = 'B,T' and @BrandID = 'ADIDAS'
begin
	INSERT INTO [dbo].[GarmentTest_Detail_Shrinkage]([ID],[No],[Location],[Type],[seq])
	select  @ID,@NO, t2.Location, t1.Type, t1.Seq 
	from GarmentTestShrinkage t1 
	inner join  #Location_S t2 on t1.Location = t2.Location
	where exists(
	select 1 from #Location_S s	
	where (
			(s.BrandID = 'ADIDAS' and t1.BrandID = s.BrandID)
			or
			(s.BrandID !='ADIDAS' and t1.BrandID = '')
		)
	)
	and t1.LocationGroup = 'TB'
end
else
begin
	INSERT INTO [dbo].[GarmentTest_Detail_Shrinkage]([ID],[No],[Location],[Type],[seq])
	select  @ID,@NO, t2.Location, t1.Type, t1.Seq 
	from GarmentTestShrinkage t1 
	inner join  #Location_S t2 on t1.LocationGroup = t2.Location
	where exists(
	select 1 from #Location_S s	
	where (
			(s.BrandID = 'ADIDAS' and t1.BrandID = s.BrandID)
			or
			(s.BrandID !='ADIDAS' and t1.BrandID = '')
		)
	)
end
INSERT INTO [dbo].[GarmentTest_Detail_Twisting]([ID],[No],[Location])
select @ID,@NO,* from #Location1
INSERT INTO [dbo].[GarmentTest_Detail_Twisting]([ID],[No],[Location])
select @ID,@NO,* from #Location2

INSERT INTO [dbo].[GarmentTest_Detail_Apperance]([ID],[No],[Type],[Seq])
values (@ID,@NO,'Printing / Heat Transfer',1)
INSERT INTO [dbo].[GarmentTest_Detail_Apperance]([ID],[No],[Type],[Seq])
values (@ID,@NO,'Label',2)
INSERT INTO [dbo].[GarmentTest_Detail_Apperance]([ID],[No],[Type],[Seq])
values (@ID,@NO,'Zipper / Snap Button / Button / Tie Cord',3)
INSERT INTO [dbo].[GarmentTest_Detail_Apperance]([ID],[No],[Type],[Seq])
values (@ID,@NO,'Discoloration (colour change )',4)
INSERT INTO [dbo].[GarmentTest_Detail_Apperance]([ID],[No],[Type],[Seq])
values (@ID,@NO,'Colour Staining',5)
INSERT INTO [dbo].[GarmentTest_Detail_Apperance]([ID],[No],[Type],[Seq])
values (@ID,@NO,'Pilling',6)
INSERT INTO [dbo].[GarmentTest_Detail_Apperance]([ID],[No],[Type],[Seq])
values (@ID,@NO,'Shrinkage & Twisting',7)
INSERT INTO [dbo].[GarmentTest_Detail_Apperance]([ID],[No],[Type],[Seq])
values (@ID,@NO,'Appearance of garment after wash',8)
";
                    ExecuteDataTableByServiceConn(CommandType.Text, insertShrinkage, objParameter1);
                    #endregion
                }

                #region 建立 Garment_Detail_Spirality
                // 代表是新增的資料
                if (dtDetail.Rows.Count == 0)
                {
                    SQLParameterCollection objParameter_Loction = new SQLParameterCollection
                    {
                        { "@StyleID", DbType.String, master.StyleID} ,
                        { "@BrandID", DbType.String, master.BrandID } ,
                        { "@SeasonID", DbType.String, master.SeasonID} ,
                    };

                    string sql_Location = @"
select sl.Location
from Style s
inner join Style_Location sl on sl.StyleUkey = s.Ukey
where s.id = @StyleID AND s.BrandID = @BrandID AND s.SeasonID = @SeasonID
";
                    DataTable dt_Location = ExecuteDataTableByServiceConn(CommandType.Text, sql_Location, objParameter_Loction);

                    string sqlcmd_Spirality = string.Empty;
                    if (dt_Location.Select("Location = 'T'").Any())
                    {
                        sqlcmd_Spirality += $@"INSERT INTO[dbo].[Garment_Detail_Spirality]([ID],[No],[Location])VALUES('{master.ID}','{item.No}','T');";
                    }

                    if (dt_Location.Select("Location = 'B'").Any())
                    {
                        sqlcmd_Spirality += $@"INSERT INTO[dbo].[Garment_Detail_Spirality]([ID],[No],[Location])VALUES('{master.ID}','{item.No}','B');";
                    }

                    ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd_Spirality, objParameter_Loction);
                }
                #endregion

                #region 寫入GarmentTest_Detail_FGPT
                SQLParameterCollection objParameter_Loctions = new SQLParameterCollection
                    {
                        { "@StyleID", DbType.String, master.StyleID} ,
                        { "@BrandID", DbType.String, master.BrandID } ,
                        { "@SeasonID", DbType.String, master.SeasonID} ,
                    };

                string sql_locations = $@"
SELECT locations = STUFF(
	(
        select DISTINCT ',' + sl.Location
	    from Style s
	    INNER JOIN Style_Location sl ON s.Ukey = sl.StyleUkey 
	    where s.id = @StyleID AND s.BrandID = @BrandID AND s.SeasonID = @SeasonID
	    FOR XML PATH('')
	) 
,1,1,'')";
                DataTable dtLocations = ExecuteDataTableByServiceConn(CommandType.Text, sql_locations, objParameter_Loctions);
                List<string> locations = dtLocations.Rows[0]["locations"].ToString().Split(',').ToList();
                bool containsT = locations.Contains("T");
                bool containsB = locations.Contains("B");

                StringBuilder insertCmd = new StringBuilder();
                List<SqlParameter> parameters = new List<SqlParameter>();
                List<GarmentTest_Detail_FGPT> fGPTs = new List<GarmentTest_Detail_FGPT>();

                string sql_RugbyFootBall = $@"select 1 from Style s where s.id = @StyleID AND s.BrandID = @BrandID AND s.SeasonID = @SeasonID AND s.ProgramID like '%FootBall%'";
                DataTable dtRugbyFootBall = ExecuteDataTableByServiceConn(CommandType.Text, sql_RugbyFootBall, objParameter_Loctions);
                bool isRugbyFootBall = dtRugbyFootBall.Rows.Count > 0;

                // 若只有B則寫入Bottom的項目+ALL的項目，若只有T則寫入TOP的項目+ALL的項目，若有B和T則寫入Top+ Bottom的項目+ALL的項目
                if (containsT && containsB)
                {
                    fGPTs = GetDefaultFGPT(false, false, true, isRugbyFootBall, "S");
                }
                else if (containsT)
                {
                    fGPTs = GetDefaultFGPT(containsT, false, false, isRugbyFootBall, "T");
                }
                else
                {
                    fGPTs = GetDefaultFGPT(false, containsB, false, isRugbyFootBall, "B");
                }

                if (item.MtlTypeID.ToString().ToUpper() == "KNIT")
                {
                    fGPTs = fGPTs.Where(w => w.TestName == "PHX-AP0450" || w.TestName == "PHX-AP0451").ToList();
                }

                int idx = 0;

                SQLParameterCollection objParameterFGPT = new SQLParameterCollection();

                foreach (var fGPT in fGPTs)
                {
                    string location = string.Empty;

                    switch (fGPT.Location)
                    {
                        case "Top":
                            location = "T";
                            break;
                        case "Bottom":
                            location = "B";
                            break;
                        case "Full": // Top+Bottom = Full
                            location = "S";
                            break;
                        default:
                            location = fGPT.Location;
                            break;
                    }

                    insertCmd.Append($@"

INSERT INTO GarmentTest_Detail_FGPT
           (ID,No,Location,Type,TestDetail,TestUnit,Criteria,TestName,Seq,TypeSelection_VersionID,TypeSelection_Seq)
     VALUES
           ( {master.ID}
           , {item.No}
           , @Location{idx}
           , @Type{idx}
           , @TestDetail{idx}
           , @TestUnit{idx}
           , @Criteria{idx}  
           , @TestName{idx}
           , '{fGPT.Seq}'
           , '{fGPT.TypeSelection_VersionID}'
           , '{fGPT.TypeSelection_Seq}')
");

                    objParameterFGPT.Add(new SqlParameter($"@Location{idx}", location));
                    objParameterFGPT.Add(new SqlParameter($"@Type{idx}", fGPT.Type));
                    objParameterFGPT.Add(new SqlParameter($"@TestDetail{idx}", fGPT.TestDetail));
                    objParameterFGPT.Add(new SqlParameter($"@TestUnit{idx}", fGPT.TestUnit));
                    objParameterFGPT.Add(new SqlParameter($"@Criteria{idx}", fGPT.Criteria));
                    objParameterFGPT.Add(new SqlParameter($"@TestName{idx}", fGPT.TestName));
                    idx++;
                }

                // 找不到才Insert
                string sql_Chk_FGPT = $"SELECT 1 FROM GarmentTest_Detail_FGPT WHERE ID ='{master.ID}' AND NO='{item.No}'";
                DataTable dtChk_FGPT = ExecuteDataTableByServiceConn(CommandType.Text, sql_Chk_FGPT, new SQLParameterCollection());
                if (dtChk_FGPT.Rows.Count == 0)
                {
                    ExecuteDataTableByServiceConn(CommandType.Text, insertCmd.ToString(), objParameterFGPT);
                }

                #endregion

                #region Save Detail 

                SQLParameterCollection objParameterDetail = new SQLParameterCollection();
                // 代表已有資料, update
                objParameterDetail.Add($"@ID", master.ID);
                objParameterDetail.Add($"@No", item.No == null ? 0 : item.No);

                objParameterDetail.Add($"@SizeCode", string.IsNullOrEmpty(item.SizeCode) ? string.Empty : item.SizeCode);
                objParameterDetail.Add($"@MtlTypeID", string.IsNullOrEmpty(item.MtlTypeID) ? string.Empty : item.MtlTypeID);
                objParameterDetail.Add($"@Result", string.IsNullOrEmpty(item.Result) ? string.Empty : item.Result);
                objParameterDetail.Add($"@NonSeamBreakageTest", item.NonSeamBreakageTest == null ? false : item.NonSeamBreakageTest);
                objParameterDetail.Add($"@SeamBreakageResult", string.IsNullOrEmpty(item.SeamBreakageResult) ? string.Empty : item.SeamBreakageResult);
                objParameterDetail.Add($"@OdourResult", string.IsNullOrEmpty(item.OdourResult) ? string.Empty : item.OdourResult);
                objParameterDetail.Add($"@WashResult", string.IsNullOrEmpty(item.WashResult) ? string.Empty : item.WashResult);
                objParameterDetail.Add($"@inspector", string.IsNullOrEmpty(item.inspector) ? string.Empty : item.inspector);                
                objParameterDetail.Add($"@Remark", string.IsNullOrEmpty(item.Remark) ? string.Empty : item.Remark);
                // objParameterDetail.Add($"@EditName", string.IsNullOrEmpty(item.EditName) ? string.Empty : item.EditName);
                // objParameterDetail.Add($"@AddName", string.IsNullOrEmpty(item.AddName) ? string.Empty : item.AddName);
                objParameterDetail.Add($"@UserID", string.IsNullOrEmpty(UserID) ? string.Empty : UserID);
                objParameterDetail.Add($"@OrderID", string.IsNullOrEmpty(item.OrderID) ? string.Empty : item.OrderID);

                //objParameterDetail.Add($"@inspdate", item.inspdate == null ? DBNull.Value : ((DateTime)item.inspdate).ToString("d"));

                string inspDate = (item.inspdate == null) ? "Null" : "'" + ((DateTime)item.inspdate).ToString("d") + "'";

                string sqlcmd = @"
update GarmentTest
set EditDate = GetDate(), EditName = @UserID
where ID = @ID
";
                if (dtDetail.Rows.Count > 0)
                {
                    sqlcmd += $@"
update GarmentTest_Detail
set SizeCode = @SizeCode
,OrderID = @OrderID
,MtlTypeID = @MtlTypeID
,Result = @Result
,NonSeamBreakageTest = @NonSeamBreakageTest
,SeamBreakageResult = @SeamBreakageResult
,OdourResult = @OdourResult
,WashResult = @WashResult
,inspector = @inspector,inspdate = {inspDate}
,Remark = @Remark
,EditName = @UserID, EditDate = GetDate()
where ID = @ID
and No = @No
" + Environment.NewLine;
                }
                else
                {
                    sqlcmd += $@"
declare @MaxNo int = (select MaxNo = isnull(max(No),0) from GarmentTest_Detail with(nolock) where  id = '{master.ID}')

insert into GarmentTest_Detail(
    ID,No,SizeCode,MtlTypeID,Result,NonSeamBreakageTest,SeamBreakageResult,OdourResult
    ,OrderID
    ,WashResult
    ,inspector
    ,inspdate
    ,Remark
    ,AddName,AddDate
    ,Status
)
values(
    @ID
    , @MaxNo +1
    , @SizeCode,@MtlTypeID,@Result,@NonSeamBreakageTest,@SeamBreakageResult,@OdourResult
    ,@OrderID
    ,@WashResult
    ,@inspector
    ,{inspDate}
    ,@Remark
    ,@UserID, GetDate()
    ,'New'
)
";
                }

                result = Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameterDetail)) > 0;

                #endregion
            }

            return result;
        }

        /// <summary>
        /// 取得預設FGPT
        /// </summary>
        /// <param name="isTop">isTop</param>
        /// <param name="isBottom">isBottom</param>
        /// <param name="isTop_Bottom">isTop_Bottom</param>
        /// <param name="isRugbyFootBall">isRugbyFootBall</param>
        /// <param name="location">location</param>
        /// <returns>List<FGPT></returns>
        public static List<GarmentTest_Detail_FGPT> GetDefaultFGPT(bool isTop, bool isBottom, bool isTop_Bottom, bool isRugbyFootBall, string location)
        {
            List<GarmentTest_Detail_FGPT> defaultFGPTList = new List<GarmentTest_Detail_FGPT>();

            List<GarmentTest_Detail_FGPT> upperOnly = new List<GarmentTest_Detail_FGPT>()
            {
                new GarmentTest_Detail_FGPT() { Seq = 3, Location = "Top", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - weft - upper bodywear 150N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 6, Location = "Top", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - warp - upper bodywear 150N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 9, Location = "Top", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - weft - upper bodywear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 12, Location = "Top", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - warp - upper bodywear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 15, Location = "Top", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - weft - upper bodywear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 18, Location = "Top", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - warp - upper bodywear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 1, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Side seam - Method B  ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 2, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Armhole seam - Method B ≥180N)", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 3, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Under arm seam or sleeve seam - Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 4, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Shoulder seam - Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 5, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Waistband seam  - Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 6, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Hood seam - Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 7, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 8, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 9, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upperr body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 19, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper  body wear/ full body wear (Side seam - Method A ≥70N)", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 20, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Armhole seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 21, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upperr body wear/ full body wear (Under arm seam or sleeve seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 22, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Shoulder seam - Method A ≥70N)", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 23, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Neck seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 24, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Waistband seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 25, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Hood seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 26, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 27, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 28, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 29, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Side seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 30, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Armhole seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 31, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Under arm seam or sleeve seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 32, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Shoulder seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 33, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Waistband seam  - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 34, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Hood seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 35, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 36, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 37, Location = "Top", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upperr body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
            };

            List<GarmentTest_Detail_FGPT> lowerOnly = new List<GarmentTest_Detail_FGPT>()
            {
                new GarmentTest_Detail_FGPT() { Seq = 1, Location = "Bottom", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - weft - lower body wear/ full body wear 150N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 4, Location = "Bottom", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - warp - lower body wear/ full body wear 150N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 7, Location = "Bottom", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - weft - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 10, Location = "Bottom", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - warp - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 13, Location = "Bottom", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - weft - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 16, Location = "Bottom", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - warp - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 10, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - length direction - lower body wear/ full body wear (Back rise- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 11, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - length direction - lower body wear/ full body wear (Crotch- Method B ≥180N)", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 12, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Front rise- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 13, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Inseam- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 14, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Sideseam- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 15, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Waistband- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 16, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 17, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 18, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 38, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Back rise- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 39, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Front rise- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 40, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Inseam- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 41, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Waistband- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 42, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Sideseam- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 43, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A  ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 44, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 45, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A  ≥70N  ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 46, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Front rise- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 47, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Back rise- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 48, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Crotch- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 49, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Inseam- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 50, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Sideseam- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 51, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Waistband- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 52, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 53, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 54, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
            };

            List<GarmentTest_Detail_FGPT> fullBody = new List<GarmentTest_Detail_FGPT>()
            {
                new GarmentTest_Detail_FGPT() { Seq = 1, Location = "Full", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - weft - lower body wear/ full body wear 150N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 4, Location = "Full", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - warp - lower body wear/ full body wear 150N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 7, Location = "Full", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - weft - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 10, Location = "Full", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - warp - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 13, Location = "Full", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - weft - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 16, Location = "Full", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - warp - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 1, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Side seam - Method B  ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 2, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Armhole seam - Method B ≥180N)", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 3, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Under arm seam or sleeve seam - Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 4, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Shoulder seam - Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 5, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Waistband seam  - Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 6, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Hood seam - Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 7, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 8, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 9, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upperr body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 10, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - length direction - lower body wear/ full body wear (Back rise- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 11, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - length direction - lower body wear/ full body wear (Crotch- Method B ≥180N)", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 12, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Front rise- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 13, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Inseam- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 14, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Sideseam- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 15, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Waistband- Method B ≥180N )", Criteria = 180 },
                new GarmentTest_Detail_FGPT() { Seq = 16, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 17, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 18, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 19, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper  body wear/ full body wear (Side seam - Method A ≥70N)", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 20, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Armhole seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 21, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upperr body wear/ full body wear (Under arm seam or sleeve seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 22, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Shoulder seam - Method A ≥70N)", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 23, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Neck seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 24, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Waistband seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 25, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Hood seam - Method A ≥70N )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 26, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 27, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 28, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 29, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Side seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 30, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Armhole seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 31, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Under arm seam or sleeve seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 32, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Shoulder seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 33, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Waistband seam  - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 34, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Hood seam - Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 35, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 36, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 37, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upperr body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 38, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Back rise- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 39, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Front rise- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 40, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Inseam- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 41, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Waistband- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 42, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Sideseam- Method A ≥70N  )", Criteria = 70 },
                new GarmentTest_Detail_FGPT() { Seq = 43, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A  ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 44, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 45, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A  ≥70N  ) Other Joining seam  selection", Criteria = 70, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 46, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Front rise- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 47, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Back rise- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 48, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Crotch- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 49, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Inseam- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 50, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Sideseam- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 51, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Waistband- Method B ≥140N )", Criteria = 140 },
                new GarmentTest_Detail_FGPT() { Seq = 52, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 53, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 54, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140, TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
            };

            List<GarmentTest_Detail_FGPT> rugby_FootBall = new List<GarmentTest_Detail_FGPT>()
            {
                new GarmentTest_Detail_FGPT() { Seq = 2, Location = "Football Style", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - weft - Rugby/Football 160N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 5, Location = "Football Style", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - warp - Rugby/Football160N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 8, Location = "Football Style", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - weft - Rugby/Football 160N", Criteria = 160 },
                new GarmentTest_Detail_FGPT() { Seq = 11, Location = "Football Style", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - warp - Rugby/Football 160N", Criteria = 160 },
                new GarmentTest_Detail_FGPT() { Seq = 14, Location = "Football Style", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - weft -Rugby/Football 160N", Criteria = 160 },
                new GarmentTest_Detail_FGPT() { Seq = 17, Location = "Football Style", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - warp - Rugby/Football 160N", Criteria = 160 },
            };

            switch (location)
            {
                case "T":
                    defaultFGPTList.AddRange(upperOnly);
                    break;
                case "B":
                    defaultFGPTList.AddRange(lowerOnly);
                    break;
                case "S":
                    defaultFGPTList.AddRange(fullBody);
                    break;
            }

            if (isRugbyFootBall)
            {
                foreach (var fGPT in rugby_FootBall)
                {
                    fGPT.Location = location;
                }

                defaultFGPTList.AddRange(rugby_FootBall);
            }

            defaultFGPTList.Add(new GarmentTest_Detail_FGPT() { Seq = 1, Location = string.Empty, Type = "odour: Garment", TestDetail = "pass/fail", TestUnit = "pass/Fail", TestName = "PHX-AP0451", });

            return defaultFGPTList.OrderBy(o => o.Type).ToList();
        }

        /*回傳Garment Test(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Garment Test
        /// </summary>
        /// <param name="Item">Garment Test成員</param>
        /// <returns>回傳Garment Test</returns>
        /// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public IList<GarmentTest_ViewModel> Get(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            StringBuilder SbSql = new StringBuilder();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,FirstOrderID"+ Environment.NewLine);
            SbSql.Append("        ,OrderID"+ Environment.NewLine);
            SbSql.Append("        ,StyleID"+ Environment.NewLine);
            SbSql.Append("        ,SeasonID"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,MDivisionid"+ Environment.NewLine);
            SbSql.Append("        ,DeadLine"+ Environment.NewLine);
            SbSql.Append("        ,SewingInline"+ Environment.NewLine);
            SbSql.Append("        ,SewingOffline"+ Environment.NewLine);
            SbSql.Append("        ,Date"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,OldUkey"+ Environment.NewLine);
            SbSql.Append("        ,SeamBreakageResult"+ Environment.NewLine);
            SbSql.Append("        ,SeamBreakageLastTestDate"+ Environment.NewLine);
            SbSql.Append("        ,OdourResult"+ Environment.NewLine);
            SbSql.Append("        ,WashResult"+ Environment.NewLine);
            SbSql.Append("FROM [GarmentTest]"+ Environment.NewLine);
            SbSql.Append("where ID = @ID" + Environment.NewLine);


            return ExecuteList<GarmentTest_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*建立Garment Test(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Garment Test
        /// </summary>
        /// <param name="Item">Garment Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Create(GarmentTest Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [GarmentTest]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
            SbSql.Append("        ,FirstOrderID"+ Environment.NewLine);
            SbSql.Append("        ,OrderID"+ Environment.NewLine);
            SbSql.Append("        ,StyleID"+ Environment.NewLine);
            SbSql.Append("        ,SeasonID"+ Environment.NewLine);
            SbSql.Append("        ,BrandID"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,MDivisionid"+ Environment.NewLine);
            SbSql.Append("        ,DeadLine"+ Environment.NewLine);
            SbSql.Append("        ,SewingInline"+ Environment.NewLine);
            SbSql.Append("        ,SewingOffline"+ Environment.NewLine);
            SbSql.Append("        ,Date"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,AddName"+ Environment.NewLine);
            SbSql.Append("        ,AddDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,OldUkey"+ Environment.NewLine);
            SbSql.Append("        ,SeamBreakageResult"+ Environment.NewLine);
            SbSql.Append("        ,SeamBreakageLastTestDate"+ Environment.NewLine);
            SbSql.Append("        ,OdourResult"+ Environment.NewLine);
            SbSql.Append("        ,WashResult"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@FirstOrderID"); objParameter.Add("@FirstOrderID", DbType.String, Item.FirstOrderID);
            SbSql.Append("        ,@OrderID"); objParameter.Add("@OrderID", DbType.String, Item.OrderID);
            SbSql.Append("        ,@StyleID"); objParameter.Add("@StyleID", DbType.String, Item.StyleID);
            SbSql.Append("        ,@SeasonID"); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
            SbSql.Append("        ,@BrandID"); objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@MDivisionid"); objParameter.Add("@MDivisionid", DbType.String, Item.MDivisionid);
            SbSql.Append("        ,@DeadLine"); objParameter.Add("@DeadLine", DbType.String, Item.DeadLine);
            SbSql.Append("        ,@SewingInline"); objParameter.Add("@SewingInline", DbType.String, Item.SewingInline);
            SbSql.Append("        ,@SewingOffline"); objParameter.Add("@SewingOffline", DbType.String, Item.SewingOffline);
            SbSql.Append("        ,@Date"); objParameter.Add("@Date", DbType.String, Item.Date);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@OldUkey"); objParameter.Add("@OldUkey", DbType.String, Item.OldUkey);
            SbSql.Append("        ,@SeamBreakageResult"); objParameter.Add("@SeamBreakageResult", DbType.String, Item.SeamBreakageResult);
            SbSql.Append("        ,@SeamBreakageLastTestDate"); objParameter.Add("@SeamBreakageLastTestDate", DbType.String, Item.SeamBreakageLastTestDate);
            SbSql.Append("        ,@OdourResult"); objParameter.Add("@OdourResult", DbType.String, Item.OdourResult);
            SbSql.Append("        ,@WashResult"); objParameter.Add("@WashResult", DbType.String, Item.WashResult);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*更新Garment Test(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Garment Test
        /// </summary>
        /// <param name="Item">Garment Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Update(GarmentTest Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [GarmentTest]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.FirstOrderID != null) { SbSql.Append(",FirstOrderID=@FirstOrderID"+ Environment.NewLine); objParameter.Add("@FirstOrderID", DbType.String, Item.FirstOrderID);}
            if (Item.OrderID != null) { SbSql.Append(",OrderID=@OrderID"+ Environment.NewLine); objParameter.Add("@OrderID", DbType.String, Item.OrderID);}
            if (Item.StyleID != null) { SbSql.Append(",StyleID=@StyleID"+ Environment.NewLine); objParameter.Add("@StyleID", DbType.String, Item.StyleID);}
            if (Item.SeasonID != null) { SbSql.Append(",SeasonID=@SeasonID"+ Environment.NewLine); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);}
            if (Item.BrandID != null) { SbSql.Append(",BrandID=@BrandID"+ Environment.NewLine); objParameter.Add("@BrandID", DbType.String, Item.BrandID);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.MDivisionid != null) { SbSql.Append(",MDivisionid=@MDivisionid"+ Environment.NewLine); objParameter.Add("@MDivisionid", DbType.String, Item.MDivisionid);}
            if (Item.DeadLine != null) { SbSql.Append(",DeadLine=@DeadLine"+ Environment.NewLine); objParameter.Add("@DeadLine", DbType.String, Item.DeadLine);}
            if (Item.SewingInline != null) { SbSql.Append(",SewingInline=@SewingInline"+ Environment.NewLine); objParameter.Add("@SewingInline", DbType.String, Item.SewingInline);}
            if (Item.SewingOffline != null) { SbSql.Append(",SewingOffline=@SewingOffline"+ Environment.NewLine); objParameter.Add("@SewingOffline", DbType.String, Item.SewingOffline);}
            if (Item.Date != null) { SbSql.Append(",Date=@Date"+ Environment.NewLine); objParameter.Add("@Date", DbType.String, Item.Date);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.OldUkey != null) { SbSql.Append(",OldUkey=@OldUkey"+ Environment.NewLine); objParameter.Add("@OldUkey", DbType.String, Item.OldUkey);}
            if (Item.SeamBreakageResult != null) { SbSql.Append(",SeamBreakageResult=@SeamBreakageResult"+ Environment.NewLine); objParameter.Add("@SeamBreakageResult", DbType.String, Item.SeamBreakageResult);}
            if (Item.SeamBreakageLastTestDate != null) { SbSql.Append(",SeamBreakageLastTestDate=@SeamBreakageLastTestDate"+ Environment.NewLine); objParameter.Add("@SeamBreakageLastTestDate", DbType.String, Item.SeamBreakageLastTestDate);}
            if (Item.OdourResult != null) { SbSql.Append(",OdourResult=@OdourResult"+ Environment.NewLine); objParameter.Add("@OdourResult", DbType.String, Item.OdourResult);}
            if (Item.WashResult != null) { SbSql.Append(",WashResult=@WashResult"+ Environment.NewLine); objParameter.Add("@WashResult", DbType.String, Item.WashResult);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		/*刪除Garment Test(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Garment Test
        /// </summary>
        /// <param name="Item">Garment Test成員</param>
        /// <returns>回傳異動筆數</returns>
		/// <info>Author: Admin; Date: 2021/08/23  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/23  1.00    Admin        Create
        /// </history>
        public int Delete(GarmentTest Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [GarmentTest]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
