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

namespace MICS.DataAccessLayer.Provider.MSSQL
{
    public class ColorFastnessProvider : SQLDAL, IColorFastnessProvider
    {
        #region 底層連線
        public ColorFastnessProvider(string ConString) : base(ConString) { }
        public ColorFastnessProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base

        public List<string> GetScales()
        {
            string sqlcmd = @"select ID from Scale  WHERE Junk=0 order by ID";
            DataTable dt = ExecuteDataTable(CommandType.Text, sqlcmd, new SQLParameterCollection());

            return dt.Rows.OfType<DataRow>().Select(dr => dr.Field<string>("ID")).ToList();
        }

        public IList<FabricColorFastness_ViewModel> GetMain(string PoID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@PoID", DbType.String, PoID } ,
            };

            string sqlcmd = @"
select a.ID,b.StyleID,b.SeasonID,b.BrandID,b.CutInLine
,a.MinSciDelivery
,a.ColorFastnessLaboratoryRemark,b.factoryid 
,[ArticlePercent] = a.LabColorFastnessPercent
,c.TestNo,c.InspDate,c.Article,c.Result
,c.Inspector,c.Remark
,[LastUpdate] = IIF(c.EditName != '',Concat((select Name from Pass1 where id = c.EditName),' ',c.EditDate)
	,Concat((select Name from Pass1 where id = c.AddName),' ',c.AddDate))
,[CompletionDate] = MaxInspDate.value
,[CreateBy] = CONCAT(a.AddName,'-',(select Name from Pass1 where id = a.AddName),' ', a.AddDate)
,[EditBy] = CONCAT(a.AddName,'-',(select Name from Pass1 where id = a.EditName),' ',a.EditDate)
from po a WITH (NOLOCK) 
left join Orders b WITH (NOLOCK) on a.ID = b.POID
left join ColorFastness c WITH (NOLOCK) on a.ID=c.POID
outer apply(
	select value = MAX(InspDate)
	from ColorFastness cf
	where cf.POID = a.ID
)MaxInspDate
where a.id= @PoID
";
            return ExecuteList<FabricColorFastness_ViewModel>(CommandType.Text, sqlcmd, objParameter);
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

            double mtlLeadT = Convert.ToDouble(ExecuteDataTableByServiceConn(CommandType.Text, "Select MtlLeadTime from System WITH (NOLOCK)", null).Rows[0]["MtlLeadTime"]);
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

        public IList<ColorFastness> Get(string ID)
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

            return ExecuteList<ColorFastness>(CommandType.Text, SbSql.ToString(), objParameter);
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
