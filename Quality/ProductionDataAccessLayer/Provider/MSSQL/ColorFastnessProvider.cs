using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using MICS.DataAccessLayer.Interface;
using DatabaseObject.ProductionDB;
using System.Linq;
using DatabaseObject.ViewModel.BulkFGT;
using System.Data.SqlClient;

namespace MICS.DataAccessLayer.Provider.MSSQL
{
    public class ColorFastnessProvider : SQLDAL, IColorFastnessProvider
    {
        #region 底層連線
        public ColorFastnessProvider(string ConString) : base(ConString) { }
        public ColorFastnessProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public bool Encode_ColorFastness(string ID, string Status, string Result, string UserID)
        {
            // 若是Amend則Result 為空白
            string strResult = (Status == "Confirmed") ? Result : "";
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
                { "@Status", DbType.String, Status } ,
                { "@result", DbType.String, strResult } ,
                { "@UserID", DbType.String, UserID } ,
            };

            string sqlcmd = @"
Update ColorFastness set Status = @Status
, result = @result
, EditName = @UserID, EditDate = GetDate()
where id = @ID

declare @POID varchar(13) = (select POID from ColorFastness WITH(NOLOCK) where ID = @ID)

exec UpdateInspPercent 'LabColorFastness', @POID
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public DataTable Get_PO_DataTable(string PoID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@POID", DbType.String, PoID },
            };

            #region Sql Command
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         ID" + Environment.NewLine);
            SbSql.Append("        ,BrandID" + Environment.NewLine);
            SbSql.Append("        ,StyleID" + Environment.NewLine);
            SbSql.Append("        ,SeasonID" + Environment.NewLine);
            SbSql.Append("FROM [PO] WITH(NOLOCK)" + Environment.NewLine);
            SbSql.Append("Where 1 = 1" + Environment.NewLine);
            #endregion

            if (!string.IsNullOrEmpty(PoID)) { SbSql.Append("And ID = @POID" + Environment.NewLine); }
            return ExecuteDataTableByServiceConn(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public DataTable Get_Mail_Content(string POID, string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@POID", DbType.String, POID } ,
                { "@ID", DbType.String, ID } ,
            };

            string sqlcmd = @"
select a.ID
,b.StyleID
,b.BrandID
,b.SeasonID
,c.TestNo
,[TestDate] = Format(c.InspDate, 'yyyy/MM/dd')
,c.Article
,c.Result
,c.Inspector
,c.Remark
from po a WITH (NOLOCK) 
left join Orders b WITH (NOLOCK) on a.ID = b.POID
left join ColorFastness c WITH (NOLOCK) on a.ID=c.POID
where a.id= @POID
and c.ID = @ID
";
            return ExecuteDataTable(CommandType.Text, sqlcmd, objParameter);
        }

        public string Get_InspectName(string Inspector)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@Inspector", DbType.String, Inspector } ,
            };

            string sqlcmd = @"
select Name from Pass1 WITH(NOLOCK) where ID = @Inspector
";
            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);
            return dt.Rows[0]["Name"].ToString();
        }

        public string Get_Supplier(string PoID, string Seq1)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@PoID", DbType.String, PoID } ,
                { "@Seq1", DbType.String, Seq1 } ,
            };

            string sqlcmd = @"
SELECT a.ID,a.SuppID,a.SEQ1
,[supplier] = a.SuppID+'-'+b.AbbEN 
from PO_Supp a WITH (NOLOCK) 
left join supp b WITH (NOLOCK) on a.SuppID=b.ID
where a.ID = @PoID
and a.seq1 = @Seq1
";
            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);
            return dt.Rows[0]["supplier"].ToString();
        }

        public List<string> GetScales()
        {
            string sqlcmd = @"select ID from Scale  WITH(NOLOCK) WHERE Junk=0 order by ID";
            DataTable dt = ExecuteDataTable(CommandType.Text, sqlcmd, new SQLParameterCollection());

            return dt.Rows.OfType<DataRow>().Select(dr => dr.Field<string>("ID")).ToList();
        }

        public FabricColorFastness_ViewModel GetMain(string PoID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@PoID", DbType.String, PoID } ,
            };

            string sqlcmd = @"
select PoID = a.ID,b.StyleID,b.SeasonID,b.BrandID
,b.CutInLine
,a.MinSciDelivery
,a.ColorFastnessLaboratoryRemark
,[ArticlePercent] = a.LabColorFastnessPercent
,[CompletionDate] = MaxInspDate.value
,[CreateBy] = CONCAT(a.AddName,'-',(select Name from Pass1 WITH(NOLOCK) where id = a.AddName),' ', a.AddDate)
,[EditBy] = CONCAT(a.AddName,'-',(select Name from Pass1 WITH(NOLOCK) where id = a.EditName),' ',a.EditDate)
from po a WITH (NOLOCK) 
left join Orders b WITH (NOLOCK) on a.ID = b.POID
outer apply(
	select value = MAX(InspDate)
	from ColorFastness cf WITH(NOLOCK)
	where cf.POID = a.ID
)MaxInspDate
where a.id= @PoID
";
            var source = ExecuteList<FabricColorFastness_ViewModel>(CommandType.Text, sqlcmd, objParameter);

            string sqlcmd2 = @"
select *,[LastUpdate] = IIF(c.EditName != '',Concat((select Name from Pass1 WITH(NOLOCK) where id = c.EditName),' ',c.EditDate)
	,Concat((select Name from Pass1 WITH(NOLOCK) where id = c.AddName),' ',c.AddDate))
from ColorFastness c WITH(NOLOCK)
where POID=@PoID
";
            var source2 = ExecuteList<ColorFastness_Result>(CommandType.Text, sqlcmd2, objParameter);

            if (source.Count == 0)
            {
                return new FabricColorFastness_ViewModel() { Result = false, ErrorMessage = "No data found" };
            }

            if (source2.Count == 0)
            {
                source2 = new List<ColorFastness_Result>();
            }

            FabricColorFastness_ViewModel result = new FabricColorFastness_ViewModel()
            {
                PoID = source.First().PoID,
                StyleID = source.First().StyleID,
                BrandID = source.First().BrandID,
                SeasonID = source.First().SeasonID,
                CutInLine = source.First().CutInLine,
                MinSciDelivery = source.First().MinSciDelivery,
                ColorFastnessLaboratoryRemark = source.First().ColorFastnessLaboratoryRemark,
                EarliestDate = source.First().CutInLine,
                EarliestSCIDel = source.First().MinSciDelivery,
                ArticlePercent = source.First().ArticlePercent,
                CompletionDate = source.First().CompletionDate,
                CreateBy = source.First().CreateBy,
                EditBy = source.First().EditBy,
                ColorFastness_MainList = source2.ToList(),
            };

            return result;
        }

        public bool Save_PO(string PoID, string Remark)
        {   
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@PoID", PoID } ,
                { "@Remark", Remark } ,
            };

            string sqlcmd = @"
update PO
set ColorFastnessLaboratoryRemark = @Remark
where ID = @PoID

exec UpdateInspPercent 'LabColorFastness',@PoID
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public bool Delete_ColorFastness(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                 { "@ID", DbType.String, ID } ,
            };
            
            string sqlcmd = $@"
SET XACT_ABORT ON

delete from ColorFastness_Detail where id = @ID
delete from ColorFastness where id = @ID
delete from [ExtendServer].PMSFile.dbo.ColorFastness where id = @ID

declare @POID varchar(13) = (select POID from ColorFastness where ID = @ID)
exec UpdateInspPercent 'LabColorFastness', @POID
";
            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public DateTime? Get_Target_LeadTime(object CUTINLINE, object MinSciDelivery)
        {
            DateTime? cutinline, sciDelv;

            if (CUTINLINE == null || string.IsNullOrEmpty(CUTINLINE.ToString()))
            {
                cutinline = null;
            }
            else
            {
                cutinline = Convert.ToDateTime(CUTINLINE);
            }

            if (MinSciDelivery == DBNull.Value || MinSciDelivery == null)
            {
                sciDelv = null;
            }
            else
            {
                sciDelv = Convert.ToDateTime(MinSciDelivery);
            }

            DateTime? targetSciDel;

            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, @"Select MtlLeadTime from Production.dbo.System WITH (NOLOCK)", new SQLParameterCollection());

            double mtlLeadT = Convert.ToDouble(dt.Rows[0]["MtlLeadTime"]);
            if (sciDelv == null)
            {
                return null;
            }

            if (string.IsNullOrEmpty(mtlLeadT.ToString()))
            {
                targetSciDel = sciDelv;
            }
            else
            {
                targetSciDel = ((DateTime)sciDelv).AddDays(Convert.ToDouble(mtlLeadT));
            }

            if (cutinline < targetSciDel)
            {
                return cutinline;
            }
            else
            {
                return targetSciDel;
            }
        }

        public IList<ColorFastness_Result> Get(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            StringBuilder SbSql = new StringBuilder();
            SbSql.Append("SELECT"+ Environment.NewLine);
            SbSql.Append("         c.ID"+ Environment.NewLine);
            SbSql.Append("        ,POID"+ Environment.NewLine);
            SbSql.Append("        ,TestNo"+ Environment.NewLine);
            SbSql.Append("        ,InspDate"+ Environment.NewLine);
            SbSql.Append("        ,Article"+ Environment.NewLine);
            SbSql.Append("        ,Result"+ Environment.NewLine);
            SbSql.Append("        ,Status"+ Environment.NewLine);
            SbSql.Append("        ,Inspector"+ Environment.NewLine);
            SbSql.Append("        ,Remark"+ Environment.NewLine);
            SbSql.Append("        ,addName"+ Environment.NewLine);
            SbSql.Append("        ,addDate"+ Environment.NewLine);
            SbSql.Append("        ,EditName"+ Environment.NewLine);
            SbSql.Append("        ,EditDate"+ Environment.NewLine);
            SbSql.Append("        ,Temperature"+ Environment.NewLine);
            SbSql.Append("        ,Cycle"+ Environment.NewLine);
            SbSql.Append("        ,Detergent"+ Environment.NewLine);
            SbSql.Append("        ,Machine"+ Environment.NewLine);
            SbSql.Append("        ,Drying"+ Environment.NewLine);
            SbSql.Append("        ,ci.TestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,ci.TestAfterPicture" + Environment.NewLine);
            SbSql.Append($@"FROM [ColorFastness] c
left join [ExtendServer].PMSFile.dbo.ColorFastness ci on c.ID=ci.ID
" + Environment.NewLine);
            SbSql.Append("where c.ID = @ID" + Environment.NewLine);

            return ExecuteList<ColorFastness_Result>(CommandType.Text, SbSql.ToString(), objParameter);
        }
       
	#endregion
    }
}
