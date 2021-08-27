using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ManufacturingExecutionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Web.Mvc;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class SearchListProvider : SQLDAL , ISearchListProvider
    {
        #region 底層連線
        public SearchListProvider(string conString) : base(conString) { }
        public SearchListProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<SelectListItem> GetTypeDatasource(string Pass1ID)
        {
            string SbSql = $@"
select Text = '', Value = ''
UNION
select Text = IIF(md.FunctionName IS NOt NULL , md.FunctionName,m.FunctionName)
	 , Value = IIF(md.FunctionName IS NOt NULL , md.FunctionName,m.FunctionName)
from Quality_Pass1 p
inner join Quality_Position pp on p.Position=pp.ID
inner join Quality_Pass2 p2 on p2.PositionID=pp.ID 
inner join Quality_Menu m on m.ID=p2.MenuID
left join Quality_Menu_detail md on md.ID=m.ID AND md.Type=p.BulkFGT_Brand
where ModuleName='Bulk FGT'
and FunctionSeq not in (10,20)
and p.ID='{Pass1ID}'
";
            return ExecuteList<SelectListItem>(CommandType.Text, SbSql, new SQLParameterCollection());
        }

        public IList<SearchList_Result> Get_SearchList(SearchList_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            StringBuilder SbSql = new StringBuilder();

            #region
            string type1 = $@"
select DISTINCT Type = 'Fabric Crocking & Shrinkage Test (504, 405)'
        , ReportNo=''
		,OrderID = o.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,Article = ''
		,Artwork = ''
		,Result=f.Result
		,TestDate = (
			SELECT MAX (TestDate) FROM (
				SELECT TestDate = f.CrockingDate
				UNION
				SELECT TestDate = f.HeatDate
				UNION 
				SELECt TestDate = f.WashDate
			)tmp
		)

from PO p
inner join Orders o ON o.POIDID= p.ID
INNER JOIN FIR_Laboratory f ON f.POID = p.ID
WHERE 1=1
";
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type1 += $@"AND o.BrandID = @BrandID";
                objParameter.Add("@BrandID", DbType.String, Req.BrandID);
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type1 += $@"AND o.SeasonID = @SeasonID";
                objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
            }
            #endregion

            #region
            string type2 = $@"
select  Type = 'Garment Test (450, 451, 701, 710)'
        ,ReportNo=gd.No
		,gd.OrderID
		,StyleID
		,BrandID
		,SeasonID
		,Article = ''
		,Artwork = ''
		,Result= IIF(gd.Result='P','Pass', IIF(gd.Result='F','Fail',''))
		,TestDate = gd.InspDate
		,g.OrderID
from GarmentTest g
inner join GarmentTest_Detail gd ON g.ID= gd.ID
WHERE 1=1
";

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type2 += ($@"AND BrandID = @BrandID");
                objParameter.Add("@BrandID", DbType.String, Req.BrandID);
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type2 += ($@"AND SeasonID = @SeasonID");
                objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
            }
            #endregion

            #region
            string type3 = $@"
select DISTINCT  Type = 'Mockup Crocking Test  (504)'
        ,ReportNo
		,OrderID = POID
		,StyleID
		,BrandID
		,SeasonID
		,Article
		,Artwork = ArtworkTypeID
		,Result
		,TestDate 
from MockupCrocking 
WHERE 1=1
";

            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                type3 += ($@"AND BrandID = @BrandID");
                objParameter.Add("@BrandID", DbType.String, Req.BrandID);
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                type3 += ($@"AND SeasonID = @SeasonID");
                objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                type3 += ($@"AND ArtworkTypeID = @Article");
                objParameter.Add("@Article", DbType.String, Req.Article);
            }
            #endregion

            // Mockup Oven Test (514)
            string type4 = $@"";
            string type5 = $@"";
            string type6 = $@"";
            string type7 = $@"";
            string type8 = $@"";
            string type9 = $@"";

            switch (Req.Type)
            {
                case "Mockup Oven Test (514)":
                    break;
                case "Mockup Wash Test (701)":
                    break;
                case "Fabric Oven Test (515)":
                    SbSql.Append($@"
select DISTINCT  ReportNo =f.TestNo
		,OrderID = o.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,Article 
		,Artwork = ''
		,Result=f.Result
		,TestDate = f.InspDate

from PO p
inner join Orders o ON o.POID = p.ID
INNER JOIN Oven f ON f.POID = p.ID
WHERE 1=1
");
                    if (!string.IsNullOrEmpty(Req.BrandID))
                    {
                        SbSql.Append($@"AND o.BrandID = @BrandID");
                        objParameter.Add("@BrandID", DbType.String, Req.BrandID);
                    }
                    if (!string.IsNullOrEmpty(Req.SeasonID))
                    {
                        SbSql.Append($@"AND o.SeasonID = @SeasonID");
                        objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
                    }
                    if (!string.IsNullOrEmpty(Req.Article))
                    {
                        SbSql.Append($@"AND Article = @Article");
                        objParameter.Add("@Article", DbType.String, Req.Article);
                    }
                    break;
                case "Fabric Color Fastness (501)":
                    SbSql.Append($@"
select DISTINCT  ReportNo =f.TestNo
		,OrderID = o.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,Article 
		,Artwork = ''
		,Result=f.Result
		,TestDate = f.InspDate

from PO p
inner join Orders o ON o.POID = p.ID
INNER JOIN ColorFastness f ON f.POID = p.ID
WHERE 1=1
");
                    if (!string.IsNullOrEmpty(Req.BrandID))
                    {
                        SbSql.Append($@"AND o.BrandID = @BrandID");
                        objParameter.Add("@BrandID", DbType.String, Req.BrandID);
                    }
                    if (!string.IsNullOrEmpty(Req.SeasonID))
                    {
                        SbSql.Append($@"AND o.SeasonID = @SeasonID");
                        objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
                    }
                    if (!string.IsNullOrEmpty(Req.Article))
                    {
                        SbSql.Append($@"AND Article = @Article");
                        objParameter.Add("@Article", DbType.String, Req.Article);
                    }
                    break;
                case "Accessory Oven Test & Wash (515, 701)":
                    SbSql.Append($@"
select DISTINCT  ReportNo=''
		,OrderID = o.POID
		,o.StyleID
		,o.BrandID
		,o.SeasonID
		,Article = ''
		,Artwork = ''
		,Result=f.Result
		,TestDate = (
			SELECT MAX (TestDate) FROM (
				SELECT TestDate = f.WashDate
				UNION
				SELECT TestDate = f.OvenDate
			)tmp
		)

from PO p
inner join Orders o ON o.POID = p.ID
INNER JOIN AIR_Laboratory f ON f.POID = p.ID
WHERE 1=1

");
                    if (!string.IsNullOrEmpty(Req.BrandID))
                    {
                        SbSql.Append($@"AND o.BrandID = @BrandID");
                        objParameter.Add("@BrandID", DbType.String, Req.BrandID);
                    }
                    if (!string.IsNullOrEmpty(Req.SeasonID))
                    {
                        SbSql.Append($@"AND o.SeasonID = @SeasonID");
                        objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
                    }
                    break;
                case "Pulling test for Snap/Botton/Rivet (437)":
                    break;
                default:
                    break;
            }

            return ExecuteList<SearchList_Result>(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
