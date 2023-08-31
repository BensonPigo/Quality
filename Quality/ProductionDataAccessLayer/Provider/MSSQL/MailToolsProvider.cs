using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.RequestModel;
using ProductionDataAccessLayer.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class MailToolsProvider : SQLDAL
    {
        #region 底層連線
        public MailToolsProvider(string ConString) : base(ConString) { }
        public MailToolsProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public string GetAIComment(SendMail_Request Req)
        {
            string sqlCmd = string.Empty;
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@Type", Req.AICommentType);

            if (Req.StyleUkey > 0)
            {
                paras.Add("@StyleUkey", Req.StyleUkey);

                sqlCmd = "select dbo.GetQualityWebAIComment(@Type,@StyleUkey,'','','')";
            }
            else if (!string.IsNullOrEmpty(Req.OrderID))
            {
                paras.Add("@OrderID", Req.OrderID);
                sqlCmd = $@"
DECLARE @StyleUkey as bigint;

select @StyleUkey = StyleUkey
from Orders 
where ID = @OrderID

select dbo.GetQualityWebAIComment(@Type,@StyleUkey,'','','')
";
            }
            else
            {
                paras.Add("@StyleID", Req.StyleID);
                paras.Add("@BrandID", Req.BrandID);
                paras.Add("@SeasonID", Req.SeasonID);

                sqlCmd = "select dbo.GetQualityWebAIComment(@Type,0,@StyleID,@BrandID,@SeasonID)";
            }

            var tmp = ExecuteScalar(System.Data.CommandType.Text, sqlCmd, paras);
            return tmp == null ? string.Empty : tmp.ToString() ;
        }
        public string GetBuyReadyDate(SendMail_Request Req)
        {
            string sqlCmd = string.Empty;
            SQLParameterCollection paras = new SQLParameterCollection();

            if (Req.StyleUkey > 0)
            {
                paras.Add("@StyleUkey", Req.StyleUkey);

                sqlCmd = $@"
Select CASE WHEN  DATEDIFF(DAY,GETDATE(),sa.BuyReadyDate) < 14 THEN 'BR Date:' + CONVERT(VARCHAR, sa.BuyReadyDate,111) + '(Less than 14 days from BR)'
			WHEN DATEDIFF(DAY,GETDATE(),sa.BuyReadyDate) >= 14 THEN 'BR Date:' + CONVERT(VARCHAR, sa.BuyReadyDate,111) + '(14 days below left for BR)'
			ELSE''
	    END
From Style_Article sa with (nolock)
inner join Style s with (nolock) on s.Ukey = sa.StyleUkey
where sa.StyleUkey = @StyleUkey
and (s.BrandID = 'ADIDAS' OR s.BrandID ='REEBOK')
";
            }
            else if (!string.IsNullOrEmpty(Req.OrderID))
            {
                paras.Add("@OrderID", Req.OrderID);
                sqlCmd = $@"
Select CASE WHEN  DATEDIFF(DAY,GETDATE(),BuyReadyDate) < 14 THEN 'BR Date:' + CONVERT(VARCHAR, BuyReadyDate,111) + '(Less than 14 days from BR)'
			WHEN DATEDIFF(DAY,GETDATE(),BuyReadyDate) >= 14 THEN 'BR Date:' + CONVERT(VARCHAR, BuyReadyDate,111) + '(14 days below left for BR)'
			ELSE''
	    END
From Style_Article sa with (nolock)
inner join Style s with (nolock) on s.Ukey = sa.StyleUkey
inner join Orders o on o.StyleUkey = s.Ukey
where o.ID = @OrderID
and (o.BrandID = 'ADIDAS' OR o.BrandID ='REEBOK')
";
            }
            else
            {
                paras.Add("@StyleID", Req.StyleID);
                paras.Add("@BrandID", Req.BrandID);
                paras.Add("@SeasonID", Req.SeasonID);

                sqlCmd = $@"
Select CASE WHEN  DATEDIFF(DAY,GETDATE(),BuyReadyDate) < 14 THEN 'BR Date:' + CONVERT(VARCHAR, BuyReadyDate,111) + '(Less than 14 days from BR)'
			WHEN DATEDIFF(DAY,GETDATE(),BuyReadyDate) >= 14 THEN 'BR Date:' + CONVERT(VARCHAR, BuyReadyDate,111) + '(14 days below left for BR)'
			ELSE''
	    END
From Style_Article sa with (nolock)
inner join Style s with (nolock) on s.Ukey = sa.StyleUkey
where s.ID = @StyleID AND s.BrandID = @BrandID AND s.SeasonID = @SeasonID
and (s.BrandID = 'ADIDAS' OR s.BrandID ='REEBOK')
";
            }

            var tmp = ExecuteScalar(System.Data.CommandType.Text, sqlCmd, paras);
            return tmp == null ? string.Empty : tmp.ToString();
        }
    }
}
