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
            SbSql.Append("FROM [PO]" + Environment.NewLine);
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
select Name from Pass1 where ID = @Inspector
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
            string sqlcmd = @"select ID from Scale  WHERE Junk=0 order by ID";
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
,[CreateBy] = CONCAT(a.AddName,'-',(select Name from Pass1 where id = a.AddName),' ', a.AddDate)
,[EditBy] = CONCAT(a.AddName,'-',(select Name from Pass1 where id = a.EditName),' ',a.EditDate)
from po a WITH (NOLOCK) 
left join Orders b WITH (NOLOCK) on a.ID = b.POID
outer apply(
	select value = MAX(InspDate)
	from ColorFastness cf
	where cf.POID = a.ID
)MaxInspDate
where a.id= @PoID
";
            var source = ExecuteList<FabricColorFastness_ViewModel>(CommandType.Text, sqlcmd, objParameter);

            string sqlcmd2 = @"
select *,[LastUpdate] = IIF(c.EditName != '',Concat((select Name from Pass1 where id = c.EditName),' ',c.EditDate)
	,Concat((select Name from Pass1 where id = c.AddName),' ',c.AddDate))
from ColorFastness c
where POID=@PoID
";
            var source2 = ExecuteList<ColorFastness_Result>(CommandType.Text, sqlcmd2, objParameter);

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

        public bool Delete_ColorFastness(string PoID, List<ColorFastness_Result> source)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            FabricColorFastness_ViewModel dbSource = GetMain(PoID);
            string sqlcmd = string.Empty;

            int idx = 1;
            foreach (var item in dbSource.ColorFastness_MainList)
            {
                if (!source.Where(x => x.ID.Equals(item.ID)).Any())
                {
                    objParameter.Add(new SqlParameter($"@ID{idx}", item.ID));
                    sqlcmd += $@"
delete from ColorFastness_Detail where id = @ID{idx} 
delete from ColorFastness where id = @ID{idx} ";
                    idx++;
                }
            }

            return Convert.ToInt32(ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter)) > 0;
        }

        public DateTime? Get_Target_LeadTime(object CUTINLINE, object MinSciDelivery)
        {
            DateTime? cutinline, sciDelv;

            if (CUTINLINE == DBNull.Value || string.IsNullOrEmpty(CUTINLINE.ToString()))
            {
                cutinline = null;
            }
            else
            {
                cutinline = Convert.ToDateTime(CUTINLINE);
            }

            if (MinSciDelivery == DBNull.Value)
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
            SbSql.Append("         ID"+ Environment.NewLine);
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
            SbSql.Append("        ,TestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture"+ Environment.NewLine);
            SbSql.Append("FROM [ColorFastness]"+ Environment.NewLine);
            SbSql.Append("where ID = @ID" + Environment.NewLine);

            return ExecuteList<ColorFastness_Result>(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Create(ColorFastness Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [ColorFastness]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID"+ Environment.NewLine);
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
            SbSql.Append("        ,TestBeforePicture"+ Environment.NewLine);
            SbSql.Append("        ,TestAfterPicture"+ Environment.NewLine);
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@POID"); objParameter.Add("@POID", DbType.String, Item.POID);
            SbSql.Append("        ,@TestNo"); objParameter.Add("@TestNo", DbType.String, Item.TestNo);
            SbSql.Append("        ,@InspDate"); objParameter.Add("@InspDate", DbType.String, Item.InspDate);
            SbSql.Append("        ,@Article"); objParameter.Add("@Article", DbType.String, Item.Article);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@Status"); objParameter.Add("@Status", DbType.String, Item.Status);
            SbSql.Append("        ,@Inspector"); objParameter.Add("@Inspector", DbType.String, Item.Inspector);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@addName"); objParameter.Add("@addName", DbType.String, Item.addName);
            SbSql.Append("        ,@addDate"); objParameter.Add("@addDate", DbType.DateTime, Item.addDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@Temperature"); objParameter.Add("@Temperature", DbType.Int32, Item.Temperature);
            SbSql.Append("        ,@Cycle"); objParameter.Add("@Cycle", DbType.Int32, Item.Cycle);
            SbSql.Append("        ,@Detergent"); objParameter.Add("@Detergent", DbType.String, Item.Detergent);
            SbSql.Append("        ,@Machine"); objParameter.Add("@Machine", DbType.String, Item.Machine);
            SbSql.Append("        ,@Drying"); objParameter.Add("@Drying", DbType.String, Item.Drying);
            SbSql.Append("        ,@TestBeforePicture"); objParameter.Add("@TestBeforePicture", DbType.String, Item.TestBeforePicture);
            SbSql.Append("        ,@TestAfterPicture"); objParameter.Add("@TestAfterPicture", DbType.String, Item.TestAfterPicture);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	
        public int Update(ColorFastness Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [ColorFastness]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.POID != null) { SbSql.Append(",POID=@POID"+ Environment.NewLine); objParameter.Add("@POID", DbType.String, Item.POID);}
            if (Item.TestNo != null) { SbSql.Append(",TestNo=@TestNo"+ Environment.NewLine); objParameter.Add("@TestNo", DbType.String, Item.TestNo);}
            if (Item.InspDate != null) { SbSql.Append(",InspDate=@InspDate"+ Environment.NewLine); objParameter.Add("@InspDate", DbType.String, Item.InspDate);}
            if (Item.Article != null) { SbSql.Append(",Article=@Article"+ Environment.NewLine); objParameter.Add("@Article", DbType.String, Item.Article);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.Status != null) { SbSql.Append(",Status=@Status"+ Environment.NewLine); objParameter.Add("@Status", DbType.String, Item.Status);}
            if (Item.Inspector != null) { SbSql.Append(",Inspector=@Inspector"+ Environment.NewLine); objParameter.Add("@Inspector", DbType.String, Item.Inspector);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.addName != null) { SbSql.Append(",addName=@addName"+ Environment.NewLine); objParameter.Add("@addName", DbType.String, Item.addName);}
            if (Item.addDate != null) { SbSql.Append(",addDate=@addDate"+ Environment.NewLine); objParameter.Add("@addDate", DbType.DateTime, Item.addDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.Temperature != null) { SbSql.Append(",Temperature=@Temperature"+ Environment.NewLine); objParameter.Add("@Temperature", DbType.Int32, Item.Temperature);}
            if (Item.Cycle != null) { SbSql.Append(",Cycle=@Cycle"+ Environment.NewLine); objParameter.Add("@Cycle", DbType.Int32, Item.Cycle);}
            if (Item.Detergent != null) { SbSql.Append(",Detergent=@Detergent"+ Environment.NewLine); objParameter.Add("@Detergent", DbType.String, Item.Detergent);}
            if (Item.Machine != null) { SbSql.Append(",Machine=@Machine"+ Environment.NewLine); objParameter.Add("@Machine", DbType.String, Item.Machine);}
            if (Item.Drying != null) { SbSql.Append(",Drying=@Drying"+ Environment.NewLine); objParameter.Add("@Drying", DbType.String, Item.Drying);}
            if (Item.TestBeforePicture != null) { SbSql.Append(",TestBeforePicture=@TestBeforePicture"+ Environment.NewLine); objParameter.Add("@TestBeforePicture", DbType.String, Item.TestBeforePicture);}
            if (Item.TestAfterPicture != null) { SbSql.Append(",TestAfterPicture=@TestAfterPicture"+ Environment.NewLine); objParameter.Add("@TestAfterPicture", DbType.String, Item.TestAfterPicture);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
		
        public int Delete(ColorFastness Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [ColorFastness]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
	#endregion
    }
}
