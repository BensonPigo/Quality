using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using MICS.DataAccessLayer.Interface;
using DatabaseObject.ProductionDB;
using DatabaseObject.ViewModel.BulkFGT;

namespace MICS.DataAccessLayer.Provider.MSSQL
{
    public class ColorFastnessDetailProvider : SQLDAL, IColorFastnessDetailProvider
    {
        #region 底層連線
        public ColorFastnessDetailProvider(string ConString) : base(ConString) { }
        public ColorFastnessDetailProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        public IList<Fabirc_ColorFastness_Detail_ViewModel> Get_DetailBody(string ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ID", DbType.String, ID } ,
            };

            string sqlcmd = @"
select cd.ID
, SubmitDate
,cd.ColorFastnessGroup
,Seq = CONCAT(cd.SEQ1,'-',cd.SEQ2)
,cd.SEQ1,cd.SEQ2
,cd.Roll
,cd.Dyelot
,po3.Refno,po3.SCIRefno,po3.ColorID
,cd.Result,cd.changeScale,cd.ResultChange,cd.StainingScale
,cd.ResultStain,cd.Remark
,[LastUpdate] = CONCAT(cd.EditName,'-',p.Name,p.ExtNo)
,cd.AddDate,cd.AddName,cd.EditDate,cd.EditName
from ColorFastness_Detail cd
left join ColorFastness c on c.ID =  cd.ID
left join PO_Supp_Detail po3 on c.POID = po3.ID 
	and cd.SEQ1 = po3.SEQ1 and cd.SEQ2 = po3.SEQ2
left join Pass1 p on p.ID = cd.EditName
where cd.ID = @ID
";
            return ExecuteList<Fabirc_ColorFastness_Detail_ViewModel>(CommandType.Text, sqlcmd, objParameter);
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
from PO_Supp_Detail
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
from FtyInventory
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

        public int Create(ColorFastness_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [ColorFastness_Detail]"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
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
            SbSql.Append(")"+ Environment.NewLine);
            SbSql.Append("VALUES"+ Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@ColorFastnessGroup"); objParameter.Add("@ColorFastnessGroup", DbType.String, Item.ColorFastnessGroup);
            SbSql.Append("        ,@SEQ1"); objParameter.Add("@SEQ1", DbType.String, Item.SEQ1);
            SbSql.Append("        ,@SEQ2"); objParameter.Add("@SEQ2", DbType.String, Item.SEQ2);
            SbSql.Append("        ,@Roll"); objParameter.Add("@Roll", DbType.String, Item.Roll);
            SbSql.Append("        ,@Dyelot"); objParameter.Add("@Dyelot", DbType.String, Item.Dyelot);
            SbSql.Append("        ,@Result"); objParameter.Add("@Result", DbType.String, Item.Result);
            SbSql.Append("        ,@changeScale"); objParameter.Add("@changeScale", DbType.String, Item.changeScale);
            SbSql.Append("        ,@StainingScale"); objParameter.Add("@StainingScale", DbType.String, Item.StainingScale);
            SbSql.Append("        ,@Remark"); objParameter.Add("@Remark", DbType.String, Item.Remark);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@SubmitDate"); objParameter.Add("@SubmitDate", DbType.String, Item.SubmitDate);
            SbSql.Append("        ,@ResultChange"); objParameter.Add("@ResultChange", DbType.String, Item.ResultChange);
            SbSql.Append("        ,@ResultStain"); objParameter.Add("@ResultStain", DbType.String, Item.ResultStain);
            SbSql.Append(")"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update(ColorFastness_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [ColorFastness_Detail]"+ Environment.NewLine);
            SbSql.Append("SET"+ Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID"+ Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID);}
            if (Item.ColorFastnessGroup != null) { SbSql.Append(",ColorFastnessGroup=@ColorFastnessGroup"+ Environment.NewLine); objParameter.Add("@ColorFastnessGroup", DbType.String, Item.ColorFastnessGroup);}
            if (Item.SEQ1 != null) { SbSql.Append(",SEQ1=@SEQ1"+ Environment.NewLine); objParameter.Add("@SEQ1", DbType.String, Item.SEQ1);}
            if (Item.SEQ2 != null) { SbSql.Append(",SEQ2=@SEQ2"+ Environment.NewLine); objParameter.Add("@SEQ2", DbType.String, Item.SEQ2);}
            if (Item.Roll != null) { SbSql.Append(",Roll=@Roll"+ Environment.NewLine); objParameter.Add("@Roll", DbType.String, Item.Roll);}
            if (Item.Dyelot != null) { SbSql.Append(",Dyelot=@Dyelot"+ Environment.NewLine); objParameter.Add("@Dyelot", DbType.String, Item.Dyelot);}
            if (Item.Result != null) { SbSql.Append(",Result=@Result"+ Environment.NewLine); objParameter.Add("@Result", DbType.String, Item.Result);}
            if (Item.changeScale != null) { SbSql.Append(",changeScale=@changeScale"+ Environment.NewLine); objParameter.Add("@changeScale", DbType.String, Item.changeScale);}
            if (Item.StainingScale != null) { SbSql.Append(",StainingScale=@StainingScale"+ Environment.NewLine); objParameter.Add("@StainingScale", DbType.String, Item.StainingScale);}
            if (Item.Remark != null) { SbSql.Append(",Remark=@Remark"+ Environment.NewLine); objParameter.Add("@Remark", DbType.String, Item.Remark);}
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName"+ Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName);}
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate"+ Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);}
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName"+ Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName);}
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate"+ Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);}
            if (Item.SubmitDate != null) { SbSql.Append(",SubmitDate=@SubmitDate"+ Environment.NewLine); objParameter.Add("@SubmitDate", DbType.String, Item.SubmitDate);}
            if (Item.ResultChange != null) { SbSql.Append(",ResultChange=@ResultChange"+ Environment.NewLine); objParameter.Add("@ResultChange", DbType.String, Item.ResultChange);}
            if (Item.ResultStain != null) { SbSql.Append(",ResultStain=@ResultStain"+ Environment.NewLine); objParameter.Add("@ResultStain", DbType.String, Item.ResultStain);}
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Delete(ColorFastness_Detail Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [ColorFastness_Detail]"+ Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
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
