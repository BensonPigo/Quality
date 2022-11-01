using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class HeatTransferWashProvider : SQLDAL
    {
        #region 底層連線
        public HeatTransferWashProvider(string ConString) : base(ConString) { }
        public HeatTransferWashProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public List<SelectListItem> GetReportNo_Source(HeatTransferWash_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select  [Text]= ReportNo, [Value]= ReportNo
from HeatTransferWash  WITH(NOLOCK)
WHERE 1=1
"; ;
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                SbSql += "AND  BrandID = @BrandID " + Environment.NewLine;
                objParameter.Add("BrandID", DbType.String, Req.BrandID);
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                SbSql += "AND  SeasonID = @SeasonID " + Environment.NewLine;
                objParameter.Add("SeasonID", DbType.String, Req.SeasonID);
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql += "AND  StyleID  = @StyleID " + Environment.NewLine;
                objParameter.Add("StyleID", DbType.String, Req.StyleID);
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                SbSql += "AND  Article = @Article " + Environment.NewLine;
                objParameter.Add("Article", DbType.String, Req.Article);
            }

            SbSql += " ORDER BY ReportNo DESC";

            return ExecuteList<SelectListItem>(CommandType.Text, SbSql, objParameter).ToList();
        }

        public HeatTransferWash_Result GetMainData(HeatTransferWash_Request Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select h.*
    ,pmsfile.TestBeforePicture
    ,pmsfile.TestAfterPicture
from HeatTransferWash h
inner join SciPMSFile_HeatTransferWash pmsfile on h.ReportNo=pmsfile.ReportNo
from HeatTransferWash h WITH(NOLOCK)
where 1 = 1 
";
            if (!string.IsNullOrEmpty(Req.ReportNo))
            {
                SbSql += "and h.ReportNo=@ReportNo" + Environment.NewLine;
                objParameter.Add("@ReportNo", DbType.String, Req.ReportNo);
            }
            if (!string.IsNullOrEmpty(Req.BrandID))
            {
                SbSql += "and h.BrandID=@BrandID" + Environment.NewLine;
                objParameter.Add("@BrandID", DbType.String, Req.BrandID);
            }
            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                SbSql += "and h.SeasonID=@SeasonID" + Environment.NewLine;
                objParameter.Add("@SeasonID", DbType.String, Req.SeasonID);
            }
            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql += "and h.StyleID=@StyleID" + Environment.NewLine;
                objParameter.Add("@StyleID", DbType.String, Req.StyleID);
            }
            if (!string.IsNullOrEmpty(Req.Article))
            {
                SbSql += "and h.Article=@Article" + Environment.NewLine;
                objParameter.Add("@Article", DbType.String, Req.Article);
            }




            List<HeatTransferWash_Result> res = ExecuteList<HeatTransferWash_Result>(CommandType.Text, SbSql, objParameter).ToList();

            return res.Any() ? res.FirstOrDefault() : new HeatTransferWash_Result();
        }
        public IList<HeatTransferWash_Detail_Result> GetDetailData(string ReportNo)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select    h.*
from HeatTransferWash_Detail h WITH(NOLOCK)
where h.ReportNo = @ReportNo
";

            objParameter.Add("ReportNo", DbType.String, ReportNo);
            IList<HeatTransferWash_Detail_Result> res = ExecuteList<HeatTransferWash_Detail_Result>(CommandType.Text, SbSql, objParameter);

            return res.Any() ? res.ToList() : new List<HeatTransferWash_Detail_Result>();
        }

        public int Insert_HeatTransferWash(HeatTransferWash_Result Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "HW", "HeatTransferWash", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, NewReportNo } ,
                { "@POID", DbType.String, Req.POID ?? ""} ,
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@Line", DbType.String,  Req.Line ?? "" } ,
                { "@Result", DbType.String,  Req.Result ?? "" } ,
                { "@Machine", DbType.String,  Req.Machine ?? "" } ,
                { "@IsTeamwear", DbType.Boolean, Req.IsTeamwear } ,
                { "@Remark", DbType.String, Req.Remark ?? "" } ,
                { "@Temperature", DbType.Int32, Req.Temperature } ,
                { "@Time", DbType.Int32, Req.Time } ,
                { "@Pressure", DbType.Decimal, Req.Pressure } ,
                { "@PeelOff", DbType.String, Req.PeelOff ?? "" } ,
                { "@Cycles", DbType.Int32, Req.Cycles } ,
                { "@TemperatureUnit", DbType.Int32, Req.TemperatureUnit } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
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
SET XACT_ABORT ON

INSERT INTO HeatTransferWash
           (ReportNo
           ,POID
           ,SeasonID
           ,BrandID
           ,StyleID
           ,Article
           ,Line
           ,Machine
           ,IsTeamwear
           ,Remark
           ,Status
           ,Temperature
           ,Time
           ,Result
           ,Pressure
           ,PeelOff
           ,Cycles
           ,TemperatureUnit
           ,AddName
           ,AddDate)
VALUES  (     
            @ReportNo
           ,@POID
           ,@StyleID
           ,@SeasonID
           ,@BrandID
           ,@Article
           ,@Line
           ,@Machine
           ,@IsTeamwear
           ,@ReportDate
           ,@Remark
           ,'New'
           ,@Temperature
           ,@Time
           ,@Result
           ,@Pressure
           ,@PeelOff
           ,@Cycles
           ,@TemperatureUnit
           ,@AddName
           ,GETDATE() )

INSERT INTO SciPMSFile_HeatTransferWash
           (ReportNo
            ,TestBeforePicture
            ,TestAfterPicture)
VALUES(    
            @ReportNo
            ,@TestBeforePicture
            ,@TestAfterPicture)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Insert_HeatTransferWash_Detail(HeatTransferWash_Detail_Result Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, Req.ReportNo } ,
                { "@FabricRefNo", DbType.String, Req.FabricRefNo ?? ""} ,
                { "@HTRefNo", DbType.String, Req.HTRefNo  ?? ""} ,
                { "@Result", DbType.String, Req.Result ?? "" } ,
                { "@Remark", DbType.String, Req.Remark ?? "" } ,
            };

            SbSql.Append($@"
INSERT INTO dbo.HeatTransferWash_Detail
           (NewReportNo
           ,FabricRefNo
           ,HTRefNo
           ,Result
           ,Remark
           ,EditName
           ,EditDate)
VALUES      (ReportNo
           ,FabricRefNo
           ,HTRefNo
           ,Result
           ,Remark
           ,EditName
           ,EditDate)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
    }
}
