using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class PublicWondowProvider : SQLDAL
    {
        #region 底層連線
        public PublicWondowProvider(string ConString) : base(ConString) { }
        public PublicWondowProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public IList<Window_Brand> Get_Brand(string ID,bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();
            SbSql.Append($@"
select ID
from Production.dbo.Brand WITH(NOLOCK)
where junk = 0

");
            if (!string.IsNullOrEmpty(ID))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND ID = @ID");
                    paras.Add("@ID", DbType.String, ID);
                }
                else
                {
                    SbSql.Append($@"AND ID LIKE @ID");
                    paras.Add("@ID", DbType.String, ID + "%");
                }
            }

            return ExecuteList<Window_Brand>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Brand> Get_Brand(string ID, string OtherTable, string OtherColumn)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();
            SbSql.Append($@"
select ID
from Production.dbo.{OtherTable} WITH(NOLOCK)
where 1=1

");
            if (!string.IsNullOrEmpty(ID))
            {

                SbSql.Append($@"AND {OtherColumn} LIKE @ID");
                paras.Add("@ID", DbType.String, ID + "%");
            }

            return ExecuteList<Window_Brand>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Season> Get_Season(string BrandID, string ID, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();
            SbSql.Append($@"
select DISTINCT ID
from Production.dbo.Season WITH(NOLOCK)
where junk = 0

");
            if (!string.IsNullOrEmpty(ID))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND ID = @ID ");
                    paras.Add("@ID", DbType.String, ID);
                }
                else
                {
                    SbSql.Append($@"AND ID LIKE @ID ");
                    paras.Add("@ID", DbType.String, ID + "%");
                }
            }

            if (!string.IsNullOrEmpty(BrandID))
            {
                SbSql.Append($@"AND BrandID = @BrandID ");
                paras.Add("@BrandID", DbType.String, BrandID);
            }

            SbSql.Append(" Order by ID desc");

            return ExecuteList<Window_Season>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Style> Get_Style(string BrandID, string SeasonID, string ID, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();
            SbSql.Append($@"
select DISTINCT ID
from Production.dbo.Style WITH(NOLOCK)
where junk = 0

");
            if (!string.IsNullOrEmpty(ID))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND ID = @ID ");
                    paras.Add("@ID", DbType.String, ID);
                }
                else
                {
                    SbSql.Append($@"AND ID LIKE @ID ");
                    paras.Add("@ID", DbType.String, ID + "%");
                }
            }

            if (!string.IsNullOrEmpty(BrandID))
            {
                SbSql.Append($@"AND BrandID = @BrandID ");
                paras.Add("@BrandID", DbType.String, BrandID);
            }
            if (!string.IsNullOrEmpty(SeasonID))
            {
                SbSql.Append($@"AND SeasonID = @SeasonID ");
                paras.Add("@SeasonID", DbType.String, SeasonID);
            }

            return ExecuteList<Window_Style>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Article> Get_Article(string OrderID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string Article, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            if (!string.IsNullOrEmpty(OrderID))
            {
                //有傳入 OrderID
                SbSql.Append($@"
select DISTINCT Article
from Production.dbo.Order_Qty WITH(NOLOCK)
where 1=1
AND ID = @OrderID
");
                paras.Add("@OrderID", DbType.String, OrderID);
            }
            else if (string.IsNullOrEmpty(OrderID))
            {
                //沒有傳入 OrderID
                SbSql.Append($@"
select DISTINCT Article
from Production.dbo.Style_Article WITH(NOLOCK)
where 1=1

");
                if (StyleUkey >0)
                {
                    SbSql.Append("AND StyleUkey = @StyleUkey ");
                    paras.Add("@StyleUkey ", DbType.Int64, StyleUkey);
                }
                if (!string.IsNullOrEmpty(StyleID) && !string.IsNullOrEmpty(BrandID) && !string.IsNullOrEmpty(SeasonID))
                {
                    SbSql.Append($@"
AND StyleUkey in (
	select Ukey
	from Production.dbo.Style WITH(NOLOCK)
	where id= @StyleID and BrandID=@BrandID and SeasonID=@SeasonID
)
");
                    paras.Add("@StyleID ", DbType.String, StyleID);
                    paras.Add("@BrandID ", DbType.String, BrandID);
                    paras.Add("@SeasonID ", DbType.String, SeasonID);
                }
            }
            else
            {
                //其餘則回傳空
                return new List<Window_Article>();
            }


            if (!string.IsNullOrEmpty(Article))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND Article = @Article ");
                    paras.Add("@Article", DbType.String, Article);
                }
                else
                {
                    SbSql.Append($@"AND Article LIKE @Article ");
                    paras.Add("@Article", DbType.String, Article + "%");
                }
            }

            return ExecuteList<Window_Article>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Article> Get_PoidArticle(string POID, Int64 StyleUkey, string StyleID, string BrandID, string SeasonID, string Article, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            if (!string.IsNullOrEmpty(POID))
            {
                //有傳入 OrderID
                SbSql.Append($@"
select DISTINCT oq.Article 
from orders o
inner join order_qty oq  ON  o.id = oq.id and oq.qty > 0 
where o.POID = @POID

");
                paras.Add("@POID", DbType.String, POID);
            }
            else if (string.IsNullOrEmpty(POID))
            {
                //沒有傳入 OrderID
                SbSql.Append($@"
select DISTINCT Article
from Production.dbo.Style_Article WITH(NOLOCK)
where 1=1

");
                if (StyleUkey > 0)
                {
                    SbSql.Append("AND StyleUkey = @StyleUkey ");
                    paras.Add("@StyleUkey ", DbType.Int64, StyleUkey);
                }
                if (!string.IsNullOrEmpty(StyleID) && !string.IsNullOrEmpty(BrandID) && !string.IsNullOrEmpty(SeasonID))
                {
                    SbSql.Append($@"
AND StyleUkey in (
	select Ukey
	from Production.dbo.Style WITH(NOLOCK)
	where id= @StyleID and BrandID=@BrandID and SeasonID=@SeasonID
)
");
                    paras.Add("@StyleID ", DbType.String, StyleID);
                    paras.Add("@BrandID ", DbType.String, BrandID);
                    paras.Add("@SeasonID ", DbType.String, SeasonID);
                }
            }
            else
            {
                //其餘則回傳空
                return new List<Window_Article>();
            }


            if (!string.IsNullOrEmpty(Article))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND Article = @Article ");
                    paras.Add("@Article", DbType.String, Article);
                }
                else
                {
                    SbSql.Append($@"AND Article LIKE @Article ");
                    paras.Add("@Article", DbType.String, Article + "%");
                }
            }

            return ExecuteList<Window_Article>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Size> Get_Size(string OrderID, Int64? StyleUkey, string BrandID, string SeasonID, string StyleID, string Article, string Size, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            if (!string.IsNullOrEmpty(OrderID))
            {
                //有傳入 OrderID
                SbSql.Append($@"
select DISTINCT SizeCode
from Production.dbo.Order_Qty WITH(NOLOCK)
where 1=1
AND ID = @OrderID
");
                paras.Add("@OrderID", DbType.String, OrderID);

                if (!string.IsNullOrEmpty(Article))
                {
                    SbSql.Append($@"AND Article = @Article ");
                    paras.Add("@Article", DbType.String, Article);
                }
            }
            else if (string.IsNullOrEmpty(OrderID))
            {
                //沒有傳入 OrderID
                SbSql.Append($@"
select DISTINCT SizeCode
from Production.dbo.Style_SizeCode WITH(NOLOCK)
where 1=1

");

                if (!string.IsNullOrEmpty(StyleID) && !string.IsNullOrEmpty(BrandID) && !string.IsNullOrEmpty(SeasonID))
                {
                    SbSql.Append($@"
AND StyleUkey in (
	select Ukey
	from Production.dbo.Style WITH(NOLOCK)
	where id= @StyleID and BrandID=@BrandID and SeasonID=@SeasonID
)
");
                    paras.Add("@StyleID ", DbType.String, StyleID);
                    paras.Add("@BrandID ", DbType.String, BrandID);
                    paras.Add("@SeasonID ", DbType.String, SeasonID);
                }
            }
            else
            {
                //其餘則回傳空
                return new List<Window_Size>();
            }


            if (!string.IsNullOrEmpty(Size))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND SizeCode = @Size ");
                    paras.Add("@Size", DbType.String, Size);
                }
                else
                {
                    SbSql.Append($@"AND SizeCode LIKE @Size ");
                    paras.Add("@Size", DbType.String, Size + "%");
                }
            }


            return ExecuteList<Window_Size>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Technician> Get_Technician(string CallFunction, string ID, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            SbSql.Append($@"
Select tch.ID
		, p1.Name
		, p1.ExtNo
		, p1.Factory
From Production.dbo.Technician tch WITH(NOLOCK)
Inner join Production.dbo.Pass1 p1 WITH(NOLOCK) on tch.ID = p1.ID
Where tch.{CallFunction} = 1

");


            if (!string.IsNullOrEmpty(ID))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND tch.ID = @ID ");
                    paras.Add("@ID", DbType.String, ID);
                }
                else
                {
                    SbSql.Append($@"AND tch.ID LIKE @ID ");
                    paras.Add("@ID", DbType.String, ID + "%");
                }
            }

            return ExecuteList<Window_Technician>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Pass1> Get_Pass1(string ID, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            SbSql.Append($@"
Select ID
		, Name
		, ExtNo
		, Factory
        , EMail
From  Pass1 p1 WITH(NOLOCK) --工廠
Where 1=1

");

            if (!string.IsNullOrEmpty(ID))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND ID = @ID ");
                    paras.Add("@ID", DbType.String, ID);
                }
                else
                {
                    SbSql.Append($@"AND ID LIKE @ID ");
                    paras.Add("@ID", DbType.String, ID + "%");
                }
            }

            return ExecuteList<Window_Pass1>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_LocalSupp> Get_LocalSupp(string SuppID, string Name, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            SbSql.Append($@"
Select ID
		, Abb
		, Name
From Production.dbo.LocalSupp  WITH(NOLOCK)-- 工廠
Where Junk = 0

");


            if (!string.IsNullOrEmpty(Name))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND Name = @Name ");
                    paras.Add("@Name", DbType.String, Name);
                }
                else
                {
                    SbSql.Append($@"AND Name LIKE @Name ");
                    paras.Add("@Name", DbType.String, Name + "%");
                }
            }

            if (!string.IsNullOrEmpty(SuppID))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND ID = @SuppID ");
                    paras.Add("@SuppID", DbType.String, SuppID);
                }
                else
                {
                    SbSql.Append($@"AND ID LIKE @SuppID ");
                    paras.Add("@SuppID", DbType.String, SuppID + "%");
                }
            }

            return ExecuteList<Window_LocalSupp>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_TPESupp> Get_TPESupp(string SuppID, string Name, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            string whereLocal = "";
            string where = "";


            if (!string.IsNullOrEmpty(Name))
            {
                if (IsExact)
                {
                    whereLocal = $@"AND Name = @Name ";
                    where = $@"AND NameEN = @Name ";

                    paras.Add("@Name", DbType.String, Name);
                }
                else
                {
                    whereLocal = $@"AND Name LIKE @Name ";
                    where = $@"AND NameEN LIKE @Name ";

                    paras.Add("@Name", DbType.String, Name + "%");
                }
            }

            if (!string.IsNullOrEmpty(SuppID))
            {
                if (IsExact)
                {
                    whereLocal = $@"AND ID = @SuppID ";
                    where = $@"AND ID = @SuppID ";

                    paras.Add("@SuppID", DbType.String, SuppID);
                }
                else
                {
                    whereLocal = $@"AND ID LIKE @SuppID ";
                    where = $@"AND ID LIKE @SuppID ";

                    paras.Add("@SuppID", DbType.String, SuppID + "%");
                }
            }


            ///工廠 用這段
            SbSql.Append($@"
Select ID
		, Abb 
		, Name
From Production.dbo.LocalSupp  WITH(NOLOCK)
Where Junk = 0
{whereLocal}
UNION
Select ID
		, Abb = AbbEN
		, Name =NameEN
From Production.dbo.Supp  WITH(NOLOCK)
Where Junk = 0
{where}
");


            return ExecuteList<Window_TPESupp>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Po_Supp_Detail> Get_Po_Supp_Detail(string POID, string FabricType, string Seq, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();
            SbSql.Append($@"
select psd.SEQ1
		, psd.SEQ2
		, psd.SCIRefno
		, psd.Refno
		, ColorID = pc.SpecValue
		, ps.SuppID
from Production.dbo.PO_Supp_Detail psd WITH(NOLOCK)
inner join Production.dbo.Po_Supp ps WITH(NOLOCK) on psd.ID = ps.ID and psd.Seq1 = ps.SEQ1
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.ID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
Where psd.FabricType = @FabricType
AND psd.ID = @POID

");
            paras.Add("@POID", DbType.String, POID);
            paras.Add("@FabricType", DbType.String, FabricType);

            if (!string.IsNullOrEmpty(Seq))
            {
                SbSql.Append($@"AND (psd.Seq1 = @Seq  OR psd.Seq2 = @Seq OR psd.Seq1+'-'+psd.Seq2 = @Seq )    ");
                paras.Add("@Seq", DbType.String, Seq);

                // IsExact沒有用到
            }

            return ExecuteList<Window_Po_Supp_Detail>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_FtyInventory> Get_FtyInventory(string POID, string Seq1, string Seq2, string Roll, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            //台北
            SbSql.Append($@"
Select DISTINCT Roll, Dyelot
From Production.dbo.FtyInventory  WITH(NOLOCK)--工廠
Where 1=1
");
            if (!string.IsNullOrEmpty(Seq1))
            {
                SbSql.Append($@"AND Seq1 = @Seq1 ");

                paras.Add("@Seq1", DbType.String, Seq1);
            }

            if (!string.IsNullOrEmpty(Seq2))
            {
                SbSql.Append($@"AND Seq2 = @Seq2 ");

                paras.Add("@Seq2", DbType.String, Seq2);
            }

            if (!string.IsNullOrEmpty(POID))
            {
                SbSql.Append($@"AND POID = @POID ");

                paras.Add("@POID", DbType.String, POID);
            }

            if (!string.IsNullOrEmpty(Roll))
            {
                SbSql.Append($@"AND Roll = @Roll ");

                paras.Add("@Roll", DbType.String, Roll);
            }

            // IsExact沒有用到

            return ExecuteList<Window_FtyInventory>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Appearance> Get_Appearance(string Lab)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            if (Lab == "1")
            {
                SbSql.Append($@"
Select ID 
     , Name 
from DropDownList  WITH(NOLOCK)
Where Type = 'Pms_LabSubProcess' 
");
            }
            else
            {
                SbSql.Append($@"
Select ID
     , Name
from DropDownList WITH(NOLOCK)
Where Type = 'Pms_LabAccessory'
");
            }

            return ExecuteList<Window_Appearance>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_SewingLine> Get_SewingLine(string FactoryID, string ID, bool IsExact)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            //台北
            SbSql.Append($@"
Select ID
From Production.dbo.SewingLine  WITH(NOLOCK)--工廠
Where Junk = 0
AND FactoryID=@FactoryID
");

            if (!string.IsNullOrEmpty(FactoryID))
            {
                SbSql.Append($@"AND FactoryID  = @FactoryID ");

                paras.Add("@FactoryID", DbType.String, FactoryID);
            }

            if (!string.IsNullOrEmpty(ID))
            {
                if (IsExact)
                {
                    SbSql.Append($@"AND ID  = @ID ");
                    paras.Add("@ID", DbType.String, ID);
                }
                else
                {
                    SbSql.Append($@"AND ID  LIKE @ID ");
                    paras.Add("@ID", DbType.String, ID + "%");

                }
            }


            return ExecuteList<Window_SewingLine>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Color> Get_Color(string BrandID, string ID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            //台北
            SbSql.Append($@"
Select DISTINCT ID, Name
From Production.dbo.Color WITH(NOLOCK) --工廠
Where Junk=0
");
            if (!string.IsNullOrEmpty(BrandID))
            {
                SbSql.Append($@"AND BrandID = @BrandID ");

                paras.Add("@BrandID", DbType.String, BrandID);
            }

            if (!string.IsNullOrEmpty(ID))
            {
                SbSql.Append($@"AND ID LIKE @ID ");

                paras.Add("@ID", DbType.String, ID + "%");
            }


            return ExecuteList<Window_Color>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_FGPT> Get_FGPT(string VersionID, string Code)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            //台北
            SbSql.Append($@"
select Seq = cast(Seq as varchar), Code 
From Production.dbo.TypeSelection WITH(NOLOCK) --工廠 
Where 1=1 
");
            if (!string.IsNullOrEmpty(VersionID))
            {
                SbSql.Append($@" AND VersionID  = @VersionID  ");

                paras.Add("@VersionID ", DbType.String, VersionID);
            }

            if (!string.IsNullOrEmpty(Code))
            {
                SbSql.Append($@" AND Code  LIKE @Code  ");

                paras.Add("@Code ", DbType.String, Code + "%");
            }

            SbSql.Append($@" ORDER BY Seq");

            return ExecuteList<Window_FGPT>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_Picture> Get_Picture(string Table, string BeforeColumn, string AfterColumn, string PKey_1, string PKey_2, string PKey_3, string PKey_4, string PKey_1_Val, string PKey_2_Val, string PKey_3_Val, string PKey_4_Val)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            string selectColumn = "";

            if (!string.IsNullOrEmpty(PKey_1))
            {
                selectColumn += $@",{PKey_1}" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(PKey_2))
            {
                selectColumn += $@",{PKey_2}" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(PKey_3))
            {
                selectColumn += $@",{PKey_3}" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(PKey_4))
            {
                selectColumn += $@",{PKey_4}" + Environment.NewLine;
            }

            //台北
            SbSql.Append($@"
select [BrforeImage]={BeforeColumn}
, [AfterImage]={AfterColumn}
{selectColumn}
From SciPMSFile_{Table}  WITH(NOLOCK)
Where 1=1
");

            if (!string.IsNullOrEmpty(PKey_1))
            {
                SbSql.Append($@"AND {PKey_1}  = @PKey_1  ");

                paras.Add("@PKey_1 ", DbType.String, PKey_1_Val);
            }

            if (!string.IsNullOrEmpty(PKey_2))
            {
                SbSql.Append($@"AND {PKey_2}  = @PKey_2  ");

                paras.Add("@PKey_2 ", DbType.String, PKey_2_Val);
            }

            if (!string.IsNullOrEmpty(PKey_3))
            {
                SbSql.Append($@"AND {PKey_3}  = @PKey_3  ");

                paras.Add("@PKey_3 ", DbType.String, PKey_3_Val);
            }

            if (!string.IsNullOrEmpty(PKey_4))
            {
                SbSql.Append($@"AND {PKey_4}  = @PKey_4  ");

                paras.Add("@PKey_4 ", DbType.String, PKey_4_Val);
            }

            return ExecuteList<Window_Picture>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_SinglePicture> Get_SinglePicture(string Table, string ColumnName, string PKey_1, string PKey_2, string PKey_3, string PKey_4, string PKey_1_Val, string PKey_2_Val, string PKey_3_Val, string PKey_4_Val)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            string selectColumn = "";

            if (!string.IsNullOrEmpty(PKey_1))
            {
                selectColumn += $@",{PKey_1}" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(PKey_2))
            {
                selectColumn += $@",{PKey_2}" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(PKey_3))
            {
                selectColumn += $@",{PKey_3}" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(PKey_4))
            {
                selectColumn += $@",{PKey_4}" + Environment.NewLine;
            }

            //台北
            SbSql.Append($@"
select [Image]={ColumnName}
{selectColumn}
From SciPMSFile_{Table}  WITH(NOLOCK)
Where 1=1
");

            if (!string.IsNullOrEmpty(PKey_1))
            {
                SbSql.Append($@"AND {PKey_1}  = @PKey_1  ");

                paras.Add("@PKey_1 ", PKey_1_Val);
            }

            if (!string.IsNullOrEmpty(PKey_2))
            {
                SbSql.Append($@"AND {PKey_2}  = @PKey_2  ");

                paras.Add("@PKey_2 ", DbType.String, PKey_2_Val);
            }

            if (!string.IsNullOrEmpty(PKey_3))
            {
                SbSql.Append($@"AND {PKey_3}  = @PKey_3  ");

                paras.Add("@PKey_3 ", DbType.String, PKey_3_Val);
            }

            if (!string.IsNullOrEmpty(PKey_4))
            {
                SbSql.Append($@"AND {PKey_4}  = @PKey_4  ");

                paras.Add("@PKey_4 ", DbType.String, PKey_4_Val);
            }

            return ExecuteList<Window_SinglePicture>(CommandType.Text, SbSql.ToString(), paras);
        }

        public IList<Window_TestFailMail> Get_TestFailMail(string FactoryID, string Type, string GroupNameList)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            //台北
            SbSql.Append($@"
select *
From Quality_MailGroup  WITH(NOLOCK)
Where 1=1
");
            if (!string.IsNullOrEmpty(FactoryID))
            {
                SbSql.Append($@"AND FactoryID  = @FactoryID  ");

                paras.Add("@FactoryID ", DbType.String, FactoryID);
            }
            if (!string.IsNullOrEmpty(Type))
            {
                SbSql.Append($@"AND Type  = @Type  ");

                paras.Add("@Type ", DbType.String, Type);
            }

            if (GroupNameList != null && GroupNameList.Any())
            {
                List<string> li = GroupNameList.Split(',').Where(o => !string.IsNullOrEmpty(o)).Distinct().ToList();
                SbSql.Append($@"AND GroupName IN ('{string.Join("','", li)}') ");
            }

            //if (!string.IsNullOrEmpty(GroupName))
            //{
            //    SbSql.Append($@"AND GroupName = @GroupName  ");

            //    paras.Add("@GroupName ", DbType.String, GroupName );
            //}



            return ExecuteList<Window_TestFailMail>(CommandType.Text, SbSql.ToString(), paras);
        }
        public IList<Window_Operation> Get_Operation(string FinalInspectionID, string Operation)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection paras = new SQLParameterCollection();

            string where = string.IsNullOrEmpty(Operation)? string.Empty :  $@" AND a.Operation LIKE @Operation";

            paras.Add("@FinalInspectionID ", DbType.String, FinalInspectionID);
            paras.Add("@Operation", DbType.String, Operation + "%");

            //台北
            SbSql.Append($@"
select distinct Operation 
INTO #base
from InlineInspection a
where EXISTS(
	select 1
	from FinalInspection_Order b
	where b.ID = @FinalInspectionID AND a.OrderID=b.OrderID
)

select Operation
	,Operator.EmployeeID --存這個
	,Operator.Name       --UI顯示這個
from #base a
outer apply(
	select  top 1 EmployeeID ,Name = FirstName+' '+LastName
	from  InlineInspection i
	inner join InlineEmployee emp on emp.EmployeeID = i.SewerID
	where EXISTS(
		select 1
		from FinalInspection_Order f
		where f.ID = @FinalInspectionID AND f.OrderID=i.OrderID
	)
    AND a.Operation=i.Operation
	order by InspectionDate desc
)Operator
where 1=1 
{where}

drop table #base
");

            return ExecuteList<Window_Operation>(CommandType.Text, SbSql.ToString(), paras);
        }
    }
}
