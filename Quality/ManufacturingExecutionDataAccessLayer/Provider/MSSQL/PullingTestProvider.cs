using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Web.Mvc;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class PullingTestProvider : SQLDAL
    {
        #region 底層連線
        public PullingTestProvider(string conString) : base(conString) { }
        public PullingTestProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public List<SelectListItem> GetReportNoList(PullingTest_ViewModel Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select  [Text]=p.ReportNo, [Value]=p.ReportNo
from PullingTest p 
WHERE 1=1
";


            //if (!string.IsNullOrEmpty(Req.ReportNo_Query))
            //{
            //    SbSql += "AND  p.ReportNo= @ReportNo " + Environment.NewLine;
            //    objParameter.Add("ReportNo", DbType.String, Req.ReportNo_Query);
            //}
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                SbSql += "AND  p.BrandID= @BrandID " + Environment.NewLine;
                objParameter.Add("BrandID", DbType.String, Req.BrandID);
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                SbSql += "AND  p.SeasonID= @SeasonID " + Environment.NewLine;
                objParameter.Add("SeasonID", DbType.String, Req.SeasonID);
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql += "AND  p.StyleID= @StyleID " + Environment.NewLine;
                objParameter.Add("StyleID", DbType.String, Req.StyleID);
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                SbSql += "AND  p.Article= @Article " + Environment.NewLine;
                objParameter.Add("Article", DbType.String, Req.Article);
            }

            SbSql += " ORDER BY p.ReportNo DESC";
            return ExecuteList<SelectListItem>(CommandType.Text, SbSql, objParameter).ToList();
        }

        public PullingTest_Result GetData(string ReportNo)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select    p.ReportNo
		,p.POID
		,p.BrandID
		,p.SeasonID
		,p.StyleID
		,p.Article
		,p.SizeCode
		,p.TestDate
		,p.Result
		,p.TestItem
		,p.PullForce
		,p.PullForceUnit
		,p.Inspector
		, InspectorName = i.Name
		,p.Time
		,p.FabricRefno
		,p.AccRefno
		,p.SnapOperator
		,p.Remark
		,p.TestBeforePicture
		,p.TestAfterPicture
		,[LastEditName]=IIF(p.EditDate IS NULL
			,(Concat (p.AddName , '-', a.Name, ' ', Cast(p.AddDate as varchar)))
			,(Concat (p.EditName, '-', e.Name, ' ', Cast(p.EditDate as varchar)))
		)
		,PullForce_Standard = s.PullForce
		,Time_Standard = s.Time
        ,TestDateText = convert(varchar, p.TestDate, 111)
from PullingTest p 
left join Production.dbo.Pass1 a ON a.ID=p.AddName 
left join Production.dbo.Pass1 e ON e.ID=p.EditName
left join Production.dbo.Pass1 i ON i.ID=p.Inspector
left join Production.dbo.Brand_PullingTestStandarList s ON s.BrandID = p.BrandID AND s.TestItem = p.TestItem AND s.PullForceUnit = p.PullForceUnit
where p.ReportNo = @ReportNo
";


            objParameter.Add("ReportNo", DbType.String, ReportNo);
            List<PullingTest_Result> res = ExecuteList<PullingTest_Result>(CommandType.Text, SbSql, objParameter).ToList();

            return res.Any() ? res.FirstOrDefault() : new PullingTest_Result();
        }

        public PullingTest_Result CheckSP(string POID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select distinct POID =  o.ID
	,o.BrandID
	,o.SeasonID
	,o.StyleID
	,Article = IIF(Article.CTN = 1 , oqq.Article, '')
from Production.dbo.Orders o 
left join Production.dbo.Order_Qty oqq ON o.id = oqq.ID
outer apply(
	select Ctn = COUNT(DISTINCT Article)
	from Production.dbo.Order_Qty oq 
	where o.id = oq.ID
)Article
where o.Category='B'
and o.Junk = 0 
and o.id = @POID
";


            objParameter.Add("POID", DbType.String, POID);
            List<PullingTest_Result> res = ExecuteList<PullingTest_Result>(CommandType.Text, SbSql, objParameter).ToList();

            return res.Any() ? res.FirstOrDefault() : new PullingTest_Result();
        }

        public PullingTest_Result GetPullForceUnit(string BrandID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select PullForceUnit = PullingTest_PullForceUnit
from Production.dbo.QABrandSetting 
where BrandID = @BrandID 
";

            objParameter.Add("BrandID", DbType.String, BrandID);
            List<PullingTest_Result> res = ExecuteList<PullingTest_Result>(CommandType.Text, SbSql, objParameter).ToList();

            return res.Any() ? res.FirstOrDefault() : new PullingTest_Result();
        }

        public PullingTest_Result GetStandard(string BrandID, string TestItem, string PullForceUnit)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select PullForce_Standard = PullForce
    ,Time_Standard = Time
from Production.dbo.Brand_PullingTestStandarList 
where BrandID = @BrandID 
AND TestItem = @TestItem 
AND PullForceUnit = @PullForceUnit 
";

            objParameter.Add("BrandID", DbType.String, BrandID);
            objParameter.Add("TestItem", DbType.String, TestItem);
            objParameter.Add("PullForceUnit", DbType.String, PullForceUnit);
            List<PullingTest_Result> res = ExecuteList<PullingTest_Result>(CommandType.Text, SbSql, objParameter).ToList();

            return res.Any() ? res.FirstOrDefault() : new PullingTest_Result();
        }

        public int Insert(PullingTest_Result Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.ReportNo +"%" } ,
                { "@POID", DbType.String, (Req.POID == null ? string.Empty : Req.POID )} ,
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@SizeCode", DbType.String,  Req.SizeCode ?? "" } ,
                { "@TestDate", DbType.DateTime, Req.TestDate.Value } ,
                { "@Result", DbType.String, Req.Result } ,
                { "@TestItem", DbType.String, Req.TestItem } ,
                { "@PullForce", DbType.Decimal, Req.PullForce } ,
                { "@PullForceUnit", DbType.String, Req.PullForceUnit ?? "" } ,
                { "@Inspector", DbType.String, Req.Inspector ?? "" } ,
                { "@Time", DbType.Int32, Req.Time } ,
                { "@FabricRefno", DbType.String, Req.FabricRefno ?? "" } ,
                { "@AccRefno", DbType.String, Req.AccRefno ?? "" } ,
                { "@SnapOperator", DbType.String, Req.SnapOperator ?? "" } ,
                { "@Remark", DbType.String, Req.Remark ?? "" } ,
                { "@AddName", DbType.String, Req.AddName ?? "" } ,
            };

            if (Req.TestBeforePicture != null)
            {
                objParameter.Add("@TestBeforePicture", Req.TestBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.TestAfterPicture != null)
            {
                objParameter.Add("@TestAfterPicture", Req.TestAfterPicture);
            }
            else
            {
                objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            SbSql.Append($@"
INSERT INTO PullingTest
           (ReportNo
           ,POID
           ,StyleID
           ,SeasonID
           ,BrandID
           ,Article
           ,SizeCode
           ,TestDate
           ,Result
           ,Inspector
           ,TestItem
           ,PullForceUnit  -- 
           ,PullForce  --
           ,Time
           ,FabricRefno
           ,AccRefno
           ,SnapOperator
           ,Remark
            ,TestBeforePicture
            ,TestAfterPicture
           ,AddDate
           ,AddName)
VALUES(    
            (   ---流水號處理
                select REPLACE( @ReportNo ,'%','') + ISNULL(REPLICATE('0',4-len( CAST( CAST( RIGHT( max(ReportNo),3) as int) + 1 as varchar) ))+ CAST( CAST( RIGHT( max(ReportNo),3) as int) + 1 as varchar),'0001')
                from PullingTest
                where ReportNo LIKE @ReportNo
            )
           ,@POID
           ,@StyleID
           ,@SeasonID
           ,@BrandID
           ,@Article
           ,@SizeCode
           ,@TestDate
           ,@Result
           ,@Inspector
           ,@TestItem
           ,@PullForceUnit
           ,@PullForce
           ,@Time
           ,@FabricRefno
           ,@AccRefno
           ,@SnapOperator
           ,@Remark
            ,@TestBeforePicture
            ,@TestAfterPicture
           ,GETDATE()
           ,@AddName)
");




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update(PullingTest_Result Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.ReportNo } ,
            };

            string modifyCol = string.Empty;

            #region 欄位SQL
            if (!string.IsNullOrEmpty(Req.Article))
            {
                modifyCol += $@"        ,Article = @Article " + Environment.NewLine;
                objParameter.Add("@Article", DbType.String, Req.Article);
            }
            if (!string.IsNullOrEmpty(Req.SizeCode))
            {
                modifyCol += $@"        ,SizeCode = @SizeCode " + Environment.NewLine;
                objParameter.Add("@SizeCode", DbType.String, Req.SizeCode);
            }
            if (Req.TestDate.HasValue)
            {
                modifyCol += $@"        ,TestDate = @TestDate " + Environment.NewLine;
                objParameter.Add("@TestDate", DbType.DateTime, Req.TestDate.Value);
            }
            else
            {
                modifyCol += $@"        ,TestDate = NULL " + Environment.NewLine;
            }
            if (!string.IsNullOrEmpty(Req.Result))
            {
                modifyCol += $@"        ,Result = @Result " + Environment.NewLine;
                objParameter.Add("@Result", DbType.String, Req.Result);
            }
            if (!string.IsNullOrEmpty(Req.TestItem))
            {
                modifyCol += $@"        ,TestItem = @TestItem " + Environment.NewLine;
                objParameter.Add("@TestItem", DbType.String, Req.TestItem);
            }

            modifyCol += $@"        ,PullForce = @PullForce " + Environment.NewLine;
            objParameter.Add("@PullForce", DbType.Decimal, Req.PullForce);

            if (!string.IsNullOrEmpty(Req.PullForceUnit))
            {
                modifyCol += $@"        ,PullForceUnit = @PullForceUnit " + Environment.NewLine;
                objParameter.Add("@PullForceUnit", DbType.String, Req.PullForceUnit ?? "");
            }

            modifyCol += $@"        ,Time = @Time " + Environment.NewLine;
            objParameter.Add("@Time", DbType.Int32, Req.Time);

            if (!string.IsNullOrEmpty(Req.FabricRefno))
            {
                modifyCol += $@"        ,FabricRefno = @FabricRefno " + Environment.NewLine;
                objParameter.Add("@FabricRefno", DbType.String, Req.FabricRefno ?? "");
            }

            if (!string.IsNullOrEmpty(Req.AccRefno))
            {
                modifyCol += $@"        ,AccRefno = @AccRefno " + Environment.NewLine;
                objParameter.Add("@AccRefno", DbType.String, Req.AccRefno ?? "");
            }
            if (!string.IsNullOrEmpty(Req.SnapOperator))
            {
                modifyCol += $@"        ,SnapOperator = @SnapOperator " + Environment.NewLine;
                objParameter.Add("@SnapOperator", DbType.String, Req.SnapOperator ?? "");
            }
            if (!string.IsNullOrEmpty(Req.Remark))
            {
                modifyCol += $@"        ,Remark = @Remark " + Environment.NewLine;
                objParameter.Add("@Remark", DbType.String, Req.Remark ?? "");
            }
            if (!string.IsNullOrEmpty(Req.EditName))
            {
                modifyCol += $@"        ,EditName = @EditName " + Environment.NewLine;
                objParameter.Add("@EditName", DbType.String, Req.EditName ?? "");
            }


            modifyCol += $@"        ,TestBeforePicture = @TestBeforePicture " + Environment.NewLine;
            modifyCol += $@"        ,TestAfterPicture = @TestAfterPicture " + Environment.NewLine;
            if (Req.TestBeforePicture != null)
            {
                objParameter.Add("@TestBeforePicture", Req.TestBeforePicture);
            }
            else
            {
                objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.TestAfterPicture != null)
            {
                objParameter.Add("@TestAfterPicture", Req.TestAfterPicture);
            }
            else
            {
                objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }
            #endregion

            SbSql.Append($@"
UPDATE PullingTest
    SET EditDate = GETDATE()
{modifyCol}
WHERE ReportNo=@ReportNo
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Delete(string ReportNo)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string modifyCol = string.Empty;


            SbSql.Append($@"
DELETE FROM PullingTest
WHERE 1=1
");
            if (!string.IsNullOrEmpty(ReportNo))
            {
                SbSql.Append($@"AND ReportNo = @ReportNo ");
                objParameter.Add("@ReportNo", DbType.String, ReportNo );
            }
            else
            {
                SbSql.Append($@"AND 1=0 ");
            }

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public DataTable GetData_DataTable(string ReportNo)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ReportNo", DbType.String, ReportNo);

            string sqlGetData = @"
select    p.ReportNo
		,p.POID
		,p.StyleID
		,p.BrandID
		,p.SeasonID
		,p.Article
		,p.SizeCode
		,p.TestDate
		,p.Result
		,p.TestItem
		,p.PullForceUnit
		,p.PullForce
		,Sec = p.Time
		, [Fabric Ref#] = p.FabricRefno
		, [Accessory Ref#] = p.AccRefno
		, [Snap Operator] = p.SnapOperator
		,p.Remark
from PullingTest p 
where p.ReportNo = @ReportNo
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);
        }
    }
}
