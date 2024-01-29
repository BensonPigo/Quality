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
using System.Transactions;
using ToolKit;
using DatabaseObject.Public;
using Sci;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class GarmentTestProvider : SQLDAL, IGarmentTestProvider
    {
        #region 底層連線
        public GarmentTestProvider(string ConString) : base(ConString) { }
        public GarmentTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<Style> GetStyleID()
        {
            string sqlcmd = @"select distinct ID from Style WITH(NOLOCK) where Junk = 0";

            return ExecuteList<Style>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }

        public IList<Brand> GetBrandID()
        {
            string sqlcmd = @"select distinct ID from Brand WITH(NOLOCK) where Junk = 0";

            return ExecuteList<Brand>(CommandType.Text, sqlcmd, new SQLParameterCollection());
        }

        public IList<Season> GetSeasonID()
        {
            string sqlcmd = @"select distinct ID from Season WITH(NOLOCK) where Junk = 0";

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
select distinct Article from GarmentTest WITH(NOLOCK)
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
,[Teamwear] =  IIF(WashName.Value is null,'','Y')
,g.Result, g.Date,g.Remark
,[GarmentTestAddName] = CONCAT(g.AddName,'-',CreatBy.Name,'',g.AddDate)
,[GarmentTestEditName] = CONCAT(g.EditName,'-',EditBy.Name,'',g.EditDate)
,g.AddName,g.EditName
from GarmentTest g WITH(NOLOCK)
left join Pass1 CreatBy WITH(NOLOCK) on CreatBy.ID = g.AddName
left join Pass1 EditBy WITH(NOLOCK) on EditBy.ID = g.EditName
outer apply(
	select MinBuyerDelivery,MinSciDelivery
	from dbo.GetSCI(g.FirstOrderID,'')
) GetSCI
outer apply(
	--說明：BrandID IN ('ADIDAS','REEBOK')，且Season為23SS之後(包含23SS)，套用UNION下面的規則，反之則套用UNION上面的規則
	select Value =  r.Name 
	from Style s WITH(NOLOCK)
	inner join Reason r WITH(NOLOCK) on s.SpecialMark = r.ID and r.ReasonTypeID = 'Style_SpecialMark'
	where s.ID = g.StyleID
	and s.BrandID = g.BrandID
	and s.SeasonID = g.SeasonID
	and (r.Name in (
                     'AMERIC. FOOT. ON-FIELD'
                    ,'AMERIC. FOOT.ON-FIELD+DISNEY'
                    ,'BASEBALL OFF FIELD'
                    ,'BASEBALL ON FIELD'
                    ,'BBALL PERFORMANCE'
                    ,'BRANDED BLANKS'
                    ,'DISNEY+BBALL PER PERFORMANCE'
                    ,'Disney+Critical Product'
                    ,'FAST TRACE+TRAINING TEAMWEAR'
                    ,'Fast Track+Critical Product'
                    ,'LACROSSE ONFIELD'
                    ,'MATCH TEAMWEAR'
                    ,'Match Teamwear+Critical P'
                    ,'NCAA ON ICE'
                    ,'NHL ON ICE'
                    ,'ON-COURT'
                    ,'SLD ON-COURT'
                    ,'SLD ON-FIELD'
                    ,'SOFTBALL ON FIELD'
                    ,'TIRO'
                    ,'Tiro+Critical Product'
                    ,'TIRO+LEGO'
                    ,'Tiro+Lego+Critical P'
                    ,'TRAINING TEAMWEAR'
                    ,'Training Teamwear+Critical P'                    
        )
    )
	AND s.SeasonID < '22ZZ'
	UNION 	
	select Value = '710' 
	from Style a WITH(NOLOCK)
	where a.ID = g.StyleID
	and a.BrandID = g.BrandID
	and a.SeasonID = g.SeasonID
	AND a.Teamwear=1
	AND a.SeasonID >= '23'
	AND a.BrandID IN ('ADIDAS','REEBOK')
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

            if (!string.IsNullOrEmpty(filter.Season))
            {
                sqlcmd += " and g.SeasonID = @SeasonID" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(filter.Style))
            {
                sqlcmd += " and g.StyleID = @StyleID" + Environment.NewLine;
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

        public string CheckInstance()
        {
            string serverName = string.Empty;

            string sql = $"SELECT ServerName = @@SERVERNAME";
            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, sql, new SQLParameterCollection());
            serverName = dt.Rows[0]["ServerName"].ToString();

            return serverName;
        }
        public void Save_GarmentTest(GarmentTest_ViewModel master, List<GarmentTest_Detail_ViewModel> detail, string UserID, bool sameInstance)
        {
            bool result = true;
            using (TransactionScope transaction = new TransactionScope())
            {
                foreach (var item in detail)
                {
                    string sql_Shrinkage_Chk = $"select 1 from Production.dbo.GarmentTest_Detail_Shrinkage with(nolock) where id = '{master.ID}' and NO = '{item.No}'";
                    DataTable dtChk_Shrinkage = ExecuteDataTableByServiceConn(CommandType.Text, sql_Shrinkage_Chk, new SQLParameterCollection());

                    string sql_detail = $@"select 1 from GarmentTest_Detail with(nolock) where  id = '{master.ID}' and NO = '{item.No}'";
                    DataTable dtDetail = ExecuteDataTableByServiceConn(CommandType.Text, sql_detail, new SQLParameterCollection());

                    if (dtChk_Shrinkage != null && dtChk_Shrinkage.Rows.Count == 0)
                    {
                        #region insertShrinkage
                        SQLParameterCollection objParameter1 = new SQLParameterCollection
                    {
                        { "@ID", DbType.String, master.ID} ,
                        { "@BrandID", DbType.String, master.BrandID.ToUpper() } ,
                        { "@No", DbType.String, item.No} ,
                    };

                        string insertShrinkage = $@"
select sl.Location
into #Location1
from GarmentTest gt with(nolock)
inner join style s with(nolock) on s.id = gt.StyleID and s.BrandID = gt.BrandID and s.SeasonID = gt.SeasonID
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
inner join style s with(nolock) on s.id = gt.StyleID and s.BrandID = gt.BrandID and s.SeasonID = gt.SeasonID
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
inner join style s with(nolock) on s.id = gt.StyleID and s.BrandID = gt.BrandID and s.SeasonID = gt.SeasonID
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

if @BrandID = 'ADIDAS' or @BrandID = 'REEBOK'
begin
	if @location_combo = 'B,T'
	begin
		INSERT INTO [dbo].[GarmentTest_Detail_Shrinkage]([ID],[No],[Location],[Type],[seq])
		select  @ID,@NO, t2.Location, t1.Type, t1.Seq
		from GarmentTestShrinkage t1  WITH(NOLOCK)
		inner join  #Location_S t2 on t1.Location = t2.Location
		where t1.BrandID = @BrandID
		and t1.LocationGroup = 'TB'
	end
	else
	begin
		INSERT INTO [dbo].[GarmentTest_Detail_Shrinkage]([ID],[No],[Location],[Type],[seq])
		select distinct  @ID,@NO, t2.Location, t1.Type, t1.Seq
		from GarmentTestShrinkage t1  WITH(NOLOCK)
		inner join  #Location_S t2 on t1.LocationGroup = t2.Location
		where t1.BrandID = @BrandID
	end
end
else
begin
	INSERT INTO [dbo].[GarmentTest_Detail_Shrinkage]([ID],[No],[Location],[Type],[seq])
	select distinct  @ID,@NO, t2.Location, t1.Type, t1.Seq
	from GarmentTestShrinkage t1  WITH(NOLOCK)
	inner join  #Location_S t2 on t1.LocationGroup = t2.Location
	where t1.BrandID = ''
end

IF NOT EXISTS (
    SELECT  1 FROM GarmentTest_Detail_Apperance
    where ID=@ID and No = @No
)
begin
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
end
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
from Style s WITH(NOLOCK)
inner join Style_Location sl WITH(NOLOCK) on sl.StyleUkey = s.Ukey
where s.id = @StyleID AND s.BrandID = @BrandID AND s.SeasonID = @SeasonID
";
                        DataTable dt_Location = ExecuteDataTableByServiceConn(CommandType.Text, sql_Location, objParameter_Loction);

                        string sqlcmd_Spirality = string.Empty;

                        if (dt_Location.Select("Location = 'T'").Any() || dt_Location.Select("Location = 'B'").Any())
                        {
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
	    from Style s WITH(NOLOCK)
	    INNER JOIN Style_Location sl WITH(NOLOCK) ON s.Ukey = sl.StyleUkey 
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

                    string sql_RugbyFootBall = $@"select 1 from Style s WITH(NOLOCK) where s.id = @StyleID AND s.BrandID = @BrandID AND s.SeasonID = @SeasonID AND s.ProgramID like '%FootBall%'";
                    DataTable dtRugbyFootBall = ExecuteDataTableByServiceConn(CommandType.Text, sql_RugbyFootBall, objParameter_Loctions);
                    bool isRugbyFootBall = dtRugbyFootBall.Rows.Count > 0 && item.MtlTypeID.ToString().ToUpper() == "WOVEN";

                    // 若只有B則寫入Bottom的項目+ALL的項目，若只有T則寫入TOP的項目+ALL的項目，若有B和T則寫入Top+ Bottom的項目+ALL的項目
                    if (containsT && containsB)
                    {
                        fGPTs = GetDefaultFGPT(false, false, true, isRugbyFootBall, "S");
                    }
                    else if (containsT)
                    {
                        fGPTs = GetDefaultFGPT(containsT, false, false, isRugbyFootBall, "T");
                    }
                    else if (containsB)
                    {
                        fGPTs = GetDefaultFGPT(false, containsB, false, isRugbyFootBall, "B");
                    }
                    else
                    {
                        fGPTs = GetDefaultFGPT(false, false, false, isRugbyFootBall, string.Empty);
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
                        objParameterFGPT.Add(new SqlParameter($"@Criteria{idx}", fGPT.Criteria ?? 0));
                        objParameterFGPT.Add(new SqlParameter($"@TestName{idx}", fGPT.TestName));
                        idx++;
                    }

                    // 找不到才Insert
                    string sql_Chk_FGPT = $"SELECT 1 FROM GarmentTest_Detail_FGPT WITH(NOLOCK) WHERE ID ='{master.ID}' AND NO='{item.No}'";
                    DataTable dtChk_FGPT = ExecuteDataTableByServiceConn(CommandType.Text, sql_Chk_FGPT, new SQLParameterCollection());
                    if (dtChk_FGPT.Rows.Count == 0)
                    {
                        ExecuteDataTableByServiceConn(CommandType.Text, insertCmd.ToString(), objParameterFGPT);
                    }

                    #endregion

                }

                #region Save Detail 

                string sqlUpdateMaster = $@"
update SampleGarmentTest
set EditDate = GetDate(), EditName = @UserID
where ID = @ID
";
                List<GarmentTest_Detail_ViewModel> oldDetailData = GetDetail(master.ID.ToString(), sameInstance).ToList();
                List<GarmentTest_Detail_ViewModel> needUpdateDetailList = PublicClass.CompareListValue<GarmentTest_Detail_ViewModel>(
                    detail, oldDetailData, "ID,No", "OrderID,SizeCode,MtlTypeID,inspdate,inspector,NonSeamBreakageTest,Remark");

                string NewReportNo = GetID(master.MDivisionid + "GM", "GarmentTest_Detail", DateTime.Today, 2, "ReportNo");

                string sqlInsertGarmentTestDetail = $@"
SET XACT_ABORT ON
declare @MaxNo int = (select MaxNo = isnull(max(No),0) from GarmentTest_Detail with(nolock) where  id = '{master.ID}')

insert into GarmentTest_Detail(
    ID,No,SizeCode,MtlTypeID,NonSeamBreakageTest
    ,OrderID
    ,inspector
    ,inspdate
    ,Remark
    ,AddName,AddDate
    ,Status
    ,ReportNo
    ,FabricationType
)
values(
    @ID
    , @MaxNo +1
    , @SizeCode,@MtlTypeID,@NonSeamBreakageTest
    ,@OrderID
    ,@inspector
    ,@inspdate
    ,@Remark
    ,@UserID, GetDate()
    ,'New'
    ,@ReportNo
    ,'Non'
)

insert into  {(sameInstance ? string.Empty : "[ExtendServer].")}PMSFile.dbo.GarmentTest_Detail(ID,No)
values(@ID, @MaxNo +1)
";
                string sqlUpdateDetail = $@"
update GarmentTest_Detail
set SizeCode = @SizeCode
,OrderID = @OrderID
,MtlTypeID = @MtlTypeID
,inspector = @inspector
,inspdate = @inspdate
,Remark = @Remark
,NonSeamBreakageTest = @NonSeamBreakageTest
,EditName = @UserID, EditDate = GetDate()
where ID = @ID
and No = @No
";
                string sqlDeleteDetail = $@"
SET XACT_ABORT ON

Delete GarmentTest_Detail_Shrinkage  where id = @ID and NO = @No
Delete GarmentTest_Detail_Apperance where id = @ID and NO = @No
Delete GarmentTest_Detail_FGWT where id = @ID and NO = @No
Delete GarmentTest_Detail_FGPT where id = @ID and NO = @No
Delete Garment_Detail_Spirality where id = @ID and NO = @No

Delete GarmentTest_Detail where id = @ID and NO = @No
Delete {(sameInstance ? string.Empty : "[ExtendServer].")}PMSFile.dbo.GarmentTest_Detail where id = @ID and NO = @No
";
                foreach (GarmentTest_Detail_ViewModel detailItem in needUpdateDetailList)
                {
                    SQLParameterCollection objParameterDetail = new SQLParameterCollection();
                    switch (detailItem.StateType)
                    {
                        case CompareStateType.Add:
                            objParameterDetail.Add($"@ID", master.ID);
                            objParameterDetail.Add($"@No", detailItem.No == null ? 0 : detailItem.No);

                            objParameterDetail.Add($"@SizeCode", string.IsNullOrEmpty(detailItem.SizeCode) ? string.Empty : detailItem.SizeCode);
                            objParameterDetail.Add($"@MtlTypeID", string.IsNullOrEmpty(detailItem.MtlTypeID) ? string.Empty : detailItem.MtlTypeID);
                            objParameterDetail.Add($"@NonSeamBreakageTest", detailItem.NonSeamBreakageTest);
                            objParameterDetail.Add($"@inspector", string.IsNullOrEmpty(detailItem.inspector) ? string.Empty : detailItem.inspector);
                            objParameterDetail.Add($"@Remark", string.IsNullOrEmpty(detailItem.Remark) ? string.Empty : detailItem.Remark);
                            objParameterDetail.Add($"@UserID", string.IsNullOrEmpty(UserID) ? string.Empty : UserID);
                            objParameterDetail.Add($"@OrderID", string.IsNullOrEmpty(detailItem.OrderID) ? string.Empty : detailItem.OrderID);
                            objParameterDetail.Add($"@inspdate", detailItem.inspdate);
                            objParameterDetail.Add($"@ReportNo", NewReportNo);

                            ExecuteNonQuery(CommandType.Text, sqlInsertGarmentTestDetail, objParameterDetail);
                            break;
                        case CompareStateType.Edit:
                            objParameterDetail.Add($"@ID", detailItem.ID);
                            objParameterDetail.Add($"@No", detailItem.No);

                            objParameterDetail.Add($"@OrderID", string.IsNullOrEmpty(detailItem.OrderID) ? "" : detailItem.OrderID);
                            objParameterDetail.Add($"@SizeCode", string.IsNullOrEmpty(detailItem.SizeCode) ? string.Empty : detailItem.SizeCode);
                            objParameterDetail.Add($"@MtlTypeID", string.IsNullOrEmpty(detailItem.MtlTypeID) ? "" : detailItem.MtlTypeID);
                            objParameterDetail.Add($"@NonSeamBreakageTest", detailItem.NonSeamBreakageTest);
                            objParameterDetail.Add($"@Remark", string.IsNullOrEmpty(detailItem.Remark) ? "" : detailItem.Remark);
                            objParameterDetail.Add($"@UserID", string.IsNullOrEmpty(UserID) ? string.Empty : UserID);
                            objParameterDetail.Add($"@inspdate", detailItem.inspdate);
                            objParameterDetail.Add($"@inspector", string.IsNullOrEmpty(detailItem.inspector) ? string.Empty : detailItem.inspector);

                            ExecuteNonQuery(CommandType.Text, sqlUpdateDetail, objParameterDetail);
                            break;
                        case CompareStateType.Delete:
                            objParameterDetail.Add($"@ID", detailItem.ID);
                            objParameterDetail.Add($"@No", detailItem.No);
                            objParameterDetail.Add($"@UserID", string.IsNullOrEmpty(UserID) ? string.Empty : UserID);

                            ExecuteNonQuery(CommandType.Text, sqlDeleteDetail, objParameterDetail);
                            break;
                        case CompareStateType.None:
                            break;
                        default:
                            break;
                    }

                    ExecuteDataTableByServiceConn(CommandType.Text, sqlUpdateMaster, objParameterDetail);
                }

                #endregion

                transaction.Complete();
            }
        }


        public void Save_New_FGPT_Item(GarmentTest_Detail_FGPT_ViewModel newItem)
        {
            using (TransactionScope transaction = new TransactionScope())
            {

                SQLParameterCollection objParameter = new SQLParameterCollection
                {
                    { "@ID", DbType.Int64, newItem.ID } ,
                    { "@No", DbType.Int32, newItem.No } ,
                    { "@Location", DbType.String, newItem.Location } ,
                    { "@Type", DbType.String, newItem.Type } ,
                    { "@TestName", DbType.String, newItem.TestName } ,
                    { "@Criteria", DbType.Int32, newItem.Criteria } ,
                    { "@TestUnit", DbType.String, newItem.TestUnit } ,
                    { "@TestDetail", DbType.String, newItem.TestDetail } ,
                    { "@IsOriginal", DbType.Boolean, true} ,
                };

                string cmd = $@"
INSERT INTO dbo.GarmentTest_Detail_FGPT
           (ID, No, Location, Type, TestName, TestDetail, Criteria, TestUnit, IsOriginal, Seq)
VALUES
           (@ID, @No, @Location, @Type, @TestName, @TestDetail, @Criteria, @TestUnit, @IsOriginal, 
                (
                    select Max(Seq) +1
                    from GarmentTest_Detail_FGPT WITH(NOLOCK)
                    where ID = @ID
                    and No = @No
                )
            )
";

                ExecuteNonQuery(CommandType.Text, cmd, objParameter);
                transaction.Complete();
            }
        }


        public void Delete_Original_FGPT_Item(GarmentTest_Detail_FGPT_ViewModel newItem)
        {
            using (TransactionScope transaction = new TransactionScope())
            {

                SQLParameterCollection objParameter = new SQLParameterCollection
                {
                    { "@ID", DbType.Int64, newItem.ID } ,
                    { "@No", DbType.Int32, newItem.No } ,
                    { "@Location", DbType.String, newItem.Location } ,
                    { "@Type", DbType.String, newItem.Type } ,
                    { "@TestName", DbType.String, newItem.TestName } ,
                    { "@Seq", DbType.Int32, newItem.Seq } ,
                };

                string cmd = $@"Delete GarmentTest_Detail_FGPT" + Environment.NewLine;

                if (newItem.ID == null || newItem.No == null || newItem.Location == null || newItem.Type == null || newItem.TestName == null || newItem.Seq == null)
                {
                    cmd += "where 1=0";
                }
                else
                {

                    cmd += "where ID=@ID AND No=@No AND Location=@Location AND Type=@Type AND TestName=@TestName AND Seq=@Seq ";
                }

                ExecuteNonQuery(CommandType.Text, cmd, objParameter);
                transaction.Complete();
            }
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
                new GarmentTest_Detail_FGPT() { Seq = 10, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - length direction - lower body wear/ full body wear (Back rise- Method B ≥180N )", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 11, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - length direction - lower body wear/ full body wear (Crotch- Method B ≥180N)", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 12, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Front rise- Method B ≥180N )", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 13, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Inseam- Method B ≥180N )", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 14, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Sideseam- Method B ≥180N )", Criteria = 180 , StandardRemark = "(only for welded/bonded seams) Method B ≥180N"},
                new GarmentTest_Detail_FGPT() { Seq = 15, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Waistband- Method B ≥180N )", Criteria = 180 , StandardRemark = "(only for welded/bonded seams) Method B ≥180N"},
                new GarmentTest_Detail_FGPT() { Seq = 16, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 17, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 18, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 38, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Back rise- Method A ≥70N  )", Criteria = 70 ,StandardRemark="Method A ≥70N" },
                new GarmentTest_Detail_FGPT() { Seq = 39, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Front rise- Method A ≥70N  )", Criteria = 70  ,StandardRemark="Method A ≥70N"},
                new GarmentTest_Detail_FGPT() { Seq = 40, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Inseam- Method A ≥70N  )", Criteria = 70  ,StandardRemark="Method A ≥70N"},
                new GarmentTest_Detail_FGPT() { Seq = 41, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Waistband- Method A ≥70N  )", Criteria = 70  ,StandardRemark="Method A ≥70N"},
                new GarmentTest_Detail_FGPT() { Seq = 42, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Sideseam- Method A ≥70N  )", Criteria = 70  ,StandardRemark="Method A ≥70N"},
                new GarmentTest_Detail_FGPT() { Seq = 43, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A  ≥70N ) Other Joining seam  selection", Criteria = 70 ,StandardRemark="Method A ≥70N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 44, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70 ,StandardRemark="Method A ≥70N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 45, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A  ≥70N  ) Other Joining seam  selection", Criteria = 70 ,StandardRemark="Method A ≥70N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 46, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Front rise- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N"},
                new GarmentTest_Detail_FGPT() { Seq = 47, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Back rise- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N"},
                new GarmentTest_Detail_FGPT() { Seq = 48, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Crotch- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N"},
                new GarmentTest_Detail_FGPT() { Seq = 49, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Inseam- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N"},
                new GarmentTest_Detail_FGPT() { Seq = 50, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Sideseam- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N"},
                new GarmentTest_Detail_FGPT() { Seq = 51, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Waistband- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N"},
                new GarmentTest_Detail_FGPT() { Seq = 52, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140,StandardRemark="Method B ≥140N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 53, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140,StandardRemark="Method B ≥140N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 54, Location = "Bottom", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140,StandardRemark="Method B ≥140N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
            };

            List<GarmentTest_Detail_FGPT> fullBody = new List<GarmentTest_Detail_FGPT>()
            {
                new GarmentTest_Detail_FGPT() { Seq = 1, Location = "Full", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - weft - lower body wear/ full body wear 150N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 4, Location = "Full", TestDetail = "mm", TestUnit = "mm", TestName = "PHX-AP0413", Type = "seam slippage: Garment - warp - lower body wear/ full body wear 150N", Criteria = 4 },
                new GarmentTest_Detail_FGPT() { Seq = 7, Location = "Full", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - weft - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 10, Location = "Full", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No fabric breakage: Garment - warp - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 13, Location = "Full", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - weft - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 16, Location = "Full", TestDetail = "pass/fail", TestUnit = "N", TestName = "PHX-AP0413", Type = "No seam breakage: Garment - warp - lower body wear/ full body wear 150N", Criteria = 150 },
                new GarmentTest_Detail_FGPT() { Seq = 1, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Side seam - Method B  ≥180N )", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 2, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Armhole seam - Method B ≥180N)", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 3, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Under arm seam or sleeve seam - Method B ≥180N )", Criteria = 180  ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N"},
                new GarmentTest_Detail_FGPT() { Seq = 4, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Shoulder seam - Method B ≥180N )", Criteria = 180  ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N"},
                new GarmentTest_Detail_FGPT() { Seq = 5, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Waistband seam  - Method B ≥180N )", Criteria = 180  ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N"},
                new GarmentTest_Detail_FGPT() { Seq = 6, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear (Hood seam - Method B ≥180N )", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 7, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N", TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 8, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N", TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 9, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - Upperr body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N", TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 10, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - length direction - lower body wear/ full body wear (Back rise- Method B ≥180N )", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 11, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - length direction - lower body wear/ full body wear (Crotch- Method B ≥180N)", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 12, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Front rise- Method B ≥180N )", Criteria = 180  ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N"},
                new GarmentTest_Detail_FGPT() { Seq = 13, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Inseam- Method B ≥180N )", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 14, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Sideseam- Method B ≥180N )", Criteria = 180  ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N"},
                new GarmentTest_Detail_FGPT() { Seq = 15, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear (Waistband- Method B ≥180N )", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N" },
                new GarmentTest_Detail_FGPT() { Seq = 16, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 17, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 18, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage after wash (only for welded/bonded seams): Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥180N ) Other Joining seam  selection", Criteria = 180 ,StandardRemark = "(only for welded/bonded seams) Method B ≥180N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 19, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper  body wear/ full body wear (Side seam - Method A ≥70N)", Criteria = 70 ,StandardRemark="Method A ≥70N"},
                new GarmentTest_Detail_FGPT() { Seq = 20, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Armhole seam - Method A ≥70N )", Criteria = 70 ,StandardRemark="Method A ≥70N"},
                new GarmentTest_Detail_FGPT() { Seq = 21, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upperr body wear/ full body wear (Under arm seam or sleeve seam - Method A ≥70N )", Criteria = 70 ,StandardRemark="Method A ≥70N"},
                new GarmentTest_Detail_FGPT() { Seq = 22, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Shoulder seam - Method A ≥70N)", Criteria = 70 ,StandardRemark="Method A ≥70N"},
                new GarmentTest_Detail_FGPT() { Seq = 23, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Neck seam - Method A ≥70N )", Criteria = 70 ,StandardRemark="Method A ≥70N"},
                new GarmentTest_Detail_FGPT() { Seq = 24, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Waistband seam - Method A ≥70N )", Criteria = 70 ,StandardRemark="Method A ≥70N" },
                new GarmentTest_Detail_FGPT() { Seq = 25, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear (Hood seam - Method A ≥70N )", Criteria = 70 ,StandardRemark="Method A ≥70N" },
                new GarmentTest_Detail_FGPT() { Seq = 26, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70 ,StandardRemark="Method A ≥70N", TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 27, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70 ,StandardRemark="Method A ≥70N", TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 28, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - Upper body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70 ,StandardRemark="Method A ≥70N", TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 29, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Side seam - Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 30, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Armhole seam - Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 31, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Under arm seam or sleeve seam - Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 32, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Shoulder seam - Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 33, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Waistband seam  - Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 34, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear (Hood seam - Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 35, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140 ,StandardRemark="Method B ≥140N", TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 36, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upper body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140 ,StandardRemark="Method B ≥140N", TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 37, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - Upperr body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140 ,StandardRemark="Method B ≥140N", TypeSelection_VersionID = 1, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 38, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Back rise- Method A ≥70N  )", Criteria = 70 ,StandardRemark="Method A ≥70N" },
                new GarmentTest_Detail_FGPT() { Seq = 39, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Front rise- Method A ≥70N  )", Criteria = 70 ,StandardRemark="Method A ≥70N" },
                new GarmentTest_Detail_FGPT() { Seq = 40, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Inseam- Method A ≥70N  )", Criteria = 70 ,StandardRemark="Method A ≥70N" },
                new GarmentTest_Detail_FGPT() { Seq = 41, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Waistband- Method A ≥70N  )", Criteria = 70 ,StandardRemark="Method A ≥70N" },
                new GarmentTest_Detail_FGPT() { Seq = 42, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Sideseam- Method A ≥70N  )", Criteria = 70 ,StandardRemark="Method A ≥70N" },
                new GarmentTest_Detail_FGPT() { Seq = 43, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A  ≥70N ) Other Joining seam  selection", Criteria = 70 ,StandardRemark="Method A ≥70N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 44, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A ≥70N ) Other Joining seam  selection", Criteria = 70 ,StandardRemark="Method A ≥70N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 45, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear ({0}- Method A  ≥70N  ) Other Joining seam  selection", Criteria = 70 ,StandardRemark="Method A ≥70N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 46, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Front rise- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 47, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Back rise- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 48, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - length direction - lower body wear/ full body wear (Crotch- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 49, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Inseam- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 50, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Sideseam- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 51, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear (Waistband- Method B ≥140N )", Criteria = 140 ,StandardRemark="Method B ≥140N" },
                new GarmentTest_Detail_FGPT() { Seq = 52, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140 ,StandardRemark="Method B ≥140N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 53, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140 ,StandardRemark="Method B ≥140N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
                new GarmentTest_Detail_FGPT() { Seq = 54, Location = "Full", TestDetail = "N", TestUnit = "N", TestName = "PHX-AP0450", Type = "seam breakage: Garment - width direction - lower body wear/ full body wear ({0}- Method B ≥140N ) Other Joining seam  selection", Criteria = 140 ,StandardRemark="Method B ≥140N", TypeSelection_VersionID = 2, TypeSelection_Seq = 1 },
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

        public bool Update_GarmentTest_Result(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            string sqlcmd = @"
--先重新判斷一次表身已Encode
UPDATE g
SET  Result =   CASE	WHEN NonSeamBreakageTest=1 AND OdourResult = 'P' AND WashResult='P'
						    THEN 'P'
						WHEN NonSeamBreakageTest=0 AND SeamBreakageResult='' 
						    THEN ''  
						WHEN NonSeamBreakageTest=0 AND SeamBreakageResult!=''  AND (SeamBreakageResult='F' OR OdourResult = 'F' OR WashResult='F') 
						    THEN 'F'  
					ELSE Result
                END
    ,EditDAte = GETDATE()
from GarmentTest_Detail g 
where id=@ID AND Status='Confirmed' ----已Encode = Confirmed

update g 
set g.SeamBreakageResult = ISNULL(SResult.SeamBreakageResult, '')
	,g.SeamBreakageLastTestDate = SResult.inspdate
	,g.OdourResult = ISNULL(All_Result.OdourResult, '')
	,g.WashResult = ISNULL(All_Result.WashResult, '')
	,g.Result = ISNULL(All_Result.Result,'')
	,g.Date = All_Result.inspdate
from GarmentTest g WITH(NOLOCK)
outer apply(
	select top 1 SeamBreakageResult,gd1.inspdate ,gd1.OdourResult,gd1.WashResult,gd1.Result
	from GarmentTest_Detail gd1 WITH(NOLOCK)
	where gd1.id=g.ID
	order by gd1.inspdate desc , gd1.EditDate desc , gd1.AddDate desc
)All_Result
outer apply(
	select top 1 SeamBreakageResult,gd1.inspdate 
	from GarmentTest_Detail gd1 WITH(NOLOCK)
	where gd1.id=g.ID and gd1.NonSeamBreakageTest = 0
	order by gd1.inspdate desc , gd1.EditDate desc , gd1.AddDate desc
) SResult
where ID = @ID
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
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
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         ID" + Environment.NewLine);
            SbSql.Append("        ,FirstOrderID" + Environment.NewLine);
            SbSql.Append("        ,OrderID" + Environment.NewLine);
            SbSql.Append("        ,StyleID" + Environment.NewLine);
            SbSql.Append("        ,SeasonID" + Environment.NewLine);
            SbSql.Append("        ,BrandID" + Environment.NewLine);
            SbSql.Append("        ,Article" + Environment.NewLine);
            SbSql.Append("        ,MDivisionid" + Environment.NewLine);
            SbSql.Append("        ,DeadLine" + Environment.NewLine);
            SbSql.Append("        ,SewingInline" + Environment.NewLine);
            SbSql.Append("        ,SewingOffline" + Environment.NewLine);
            SbSql.Append("        ,Date" + Environment.NewLine);
            SbSql.Append("        ,Result" + Environment.NewLine);
            SbSql.Append("        ,Remark" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append("        ,OldUkey" + Environment.NewLine);
            SbSql.Append("        ,SeamBreakageResult" + Environment.NewLine);
            SbSql.Append("        ,SeamBreakageLastTestDate" + Environment.NewLine);
            SbSql.Append("        ,OdourResult" + Environment.NewLine);
            SbSql.Append("        ,WashResult" + Environment.NewLine);
            SbSql.Append("FROM [GarmentTest] WITH(NOLOCK)" + Environment.NewLine);
            SbSql.Append("where ID = @ID" + Environment.NewLine);


            return ExecuteList<GarmentTest_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<GarmentTest_Detail_ViewModel> GetDetail(string ID, bool sameInstance)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         gd.ID" + Environment.NewLine);
            SbSql.Append("        ,gd.No" + Environment.NewLine);
            SbSql.Append("        ,Result" + Environment.NewLine);
            SbSql.Append("        ,inspdate" + Environment.NewLine);
            SbSql.Append("        ,inspector" + Environment.NewLine);
            SbSql.Append("        ,Remark" + Environment.NewLine);
            SbSql.Append("        ,Sender" + Environment.NewLine);
            SbSql.Append("        ,SendDate" + Environment.NewLine);
            SbSql.Append("        ,Receiver" + Environment.NewLine);
            SbSql.Append("        ,ReceiveDate" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append("        ,OldUkey" + Environment.NewLine);
            SbSql.Append("        ,SubmitDate" + Environment.NewLine);
            SbSql.Append("        ,ArrivedQty" + Environment.NewLine);
            SbSql.Append("        ,LineDry" + Environment.NewLine);
            SbSql.Append("        ,Temperature" + Environment.NewLine);
            SbSql.Append("        ,TumbleDry" + Environment.NewLine);
            SbSql.Append("        ,Machine" + Environment.NewLine);
            SbSql.Append("        ,HandWash" + Environment.NewLine);
            SbSql.Append("        ,Composition" + Environment.NewLine);
            SbSql.Append("        ,Neck" + Environment.NewLine);
            SbSql.Append("        ,Status" + Environment.NewLine);
            SbSql.Append("        ,SizeCode" + Environment.NewLine);
            SbSql.Append("        ,LOtoFactory" + Environment.NewLine);
            SbSql.Append("        ,MtlTypeID" + Environment.NewLine);
            SbSql.Append("        ,Above50NaturalFibres" + Environment.NewLine);
            SbSql.Append("        ,Above50SyntheticFibres" + Environment.NewLine);
            SbSql.Append("        ,OrderID" + Environment.NewLine);
            SbSql.Append("        ,NonSeamBreakageTest" + Environment.NewLine);
            SbSql.Append("        ,SeamBreakageResult" + Environment.NewLine);
            SbSql.Append("        ,OdourResult" + Environment.NewLine);
            SbSql.Append("        ,WashResult" + Environment.NewLine);
            SbSql.Append("        ,gdi.TestBeforePicture" + Environment.NewLine);
            SbSql.Append("        ,gdi.TestAfterPicture" + Environment.NewLine);
            SbSql.Append($@"FROM [GarmentTest_Detail] gd WITH(NOLOCK)
left join {(sameInstance ? string.Empty : "[ExtendServer].")}PMSFile.dbo.GarmentTest_Detail gdi WITH(NOLOCK) on gd.ID=gdi.ID AND gd.No = gdi.No
" + Environment.NewLine);
            SbSql.Append("where gd.ID = @ID" + Environment.NewLine);

            return ExecuteList<GarmentTest_Detail_ViewModel>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public IList<GarmentTest_Detail_ViewModel> GetDetail_LastTestNo(GarmentTest_Request filter, string Type)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@BrandID", DbType.String, filter.Brand } ,
                { "@StyleID", DbType.String, filter.Style } ,
                { "@SeasonID", DbType.String, filter.Season } ,
                { "@Article", DbType.String, filter.Article },
            };

            string sql;
            switch (Type)
            {
                case "450":
                    sql = @"
select *
from GarmentTest_Detail gd WITH(NOLOCK)
inner join (
	select gd2.ID, NO = MAX(gd2.NO)
	from GarmentTest_Detail gd2 WITH(NOLOCK)
	where gd2.ID in (
		select MAX(ID)
		from GarmentTest g WITH(NOLOCK)
		where g.BrandID = @BrandID
		and g.StyleID = @StyleID
		and g.SeasonID = @SeasonID
		and g.SeamBreakageLastTestDate in (
			select MAX(SeamBreakageLastTestDate)
			from GarmentTest g2 WITH(NOLOCK)
			where g2.StyleID = g.StyleID
			and g2.BrandID = g.BrandID
			and g2.SeasonID = g.SeasonID
			and g2.SeamBreakageLastTestDate is not null
		))
	and isnull(gd2.SeamBreakageResult, '') <> ''
	group by gd2.ID
) gd2 on gd.ID = gd2.ID and gd.No = gd2.NO";
                    break;
                case "451":
                    sql = @"
select gd.ID, [No] = MAX(gd.No)
from GarmentTest g WITH(NOLOCK)
inner join GarmentTest_Detail gd WITH(NOLOCK) on g.ID = gd.ID
inner join (
	select gd.ID, gd.No, date = MAX(g.date)
	from GarmentTest g WITH(NOLOCK)
	inner join GarmentTest_Detail gd WITH(NOLOCK) on g.ID = gd.ID
	where g.BrandID = @BrandID
	and g.StyleID = @StyleID
	and g.SeasonID = @SeasonID
	and g.Article = @Article
	and gd.OdourResult <> ''
	group by gd.ID, gd.No
)gd2 on gd.ID = gd2.ID and gd.No = gd2.No and g.date = gd2.date
Group by gd.ID
";
                    break;
                default:
                    sql = @"
select gd.ID, [No] = MAX(gd.No)
from GarmentTest g WITH(NOLOCK)
inner join GarmentTest_Detail gd WITH(NOLOCK) on g.ID = gd.ID
inner join (
	select gd.ID, gd.No, date = MAX(g.date)
	from GarmentTest g WITH(NOLOCK)
	inner join GarmentTest_Detail gd WITH(NOLOCK) on g.ID = gd.ID
	where g.BrandID = @BrandID
	and g.StyleID = @StyleID
	and g.SeasonID = @SeasonID
	and g.Article = @Article
	and gd.WashResult <> ''
	group by gd.ID, gd.No
)gd2 on gd.ID = gd2.ID and gd.No = gd2.No and g.date = gd2.date
Group by gd.ID
";
                    break;
            }

            return ExecuteList<GarmentTest_Detail_ViewModel>(CommandType.Text, sql, objParameter);
        }
    }
}
