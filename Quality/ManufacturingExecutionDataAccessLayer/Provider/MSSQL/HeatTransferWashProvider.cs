using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using ToolKit;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class HeatTransferWashProvider : SQLDAL
    {
        public string GetFactoryNameEN(string ReportNo, string FactoryID)
        {
            string factoryNameEN = string.Empty;
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@ReportNo", DbType.String, ReportNo } ,
                { "@FactoryID",DbType.String, FactoryID } ,
            };
            string sql = $@"
            SELECT
            o.FactoryID
            INTO #tmp
            FROM HeatTransferWash P WITH(NOLOCK)
            INNER JOIN Production.dbo.Orders O WITH(NOLOCK) ON O.ID = P.OrderID
            WHERE P.ReportNo = @ReportNo
			
            SELECT
            F.NameEN,*
            FROM Production.dbo.Factory F WITH(NOLOCK)
            LEFT JOIN #TMP T WITH(NOLOCK) ON T.FactoryID = F.ID
            WHERE F.ID = IIF((SELECT count(1) from #tmp) > 0 ,T.FactoryID,@FactoryID)";

            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, sql, objParameter);
            factoryNameEN = dt.Rows[0]["NameEN"].ToString();
            return factoryNameEN;
        }

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
select h.ReportNo
      ,h.OrderID
      ,h.BrandID
      ,h.SeasonID
      ,h.StyleID
      ,h.Article
      ,h.Line
      ,h.Machine
      ,s.Teamwear
      ,h.ReportDate
      ,h.Result
      ,h.Remark
      ,h.Status
      ,h.ArtworkTypeID
      ,h.AddName
      ,h.AddDate
      ,h.EditName
      ,h.EditDate
      ,TestBeforePicture = (select top 1 TestBeforePicture from SciPMSFile_HeatTransferWash pmsfile WITH(NOLOCK) where h.ReportNo=pmsfile.ReportNo )
      ,TestAfterPicture = (select top 1 TestAfterPicture from SciPMSFile_HeatTransferWash pmsfile WITH(NOLOCK) where h.ReportNo=pmsfile.ReportNo )
	  ,LastEditText = ISNULL(LastEdit2.Name, LastEdit1.Name)
      ,MRHandleEmail = ISNULL(p.Email, p2.Email)
      ,Signature = Technician.Signature
      ,ArtworkTypeID_FullName = SubProcessInfo.ID
      ,h.ReceivedDate
      ,[Preparer] = h.Preparer
      ,[PreparerName] = (select Name from pass1 where id = h.Preparer)
      ,h.Approver
      ,[ApproverName] = pass1Approver.Name
      ,ApproverSignature = ApproverSignature.Signature
      ,h.Inspector
      ,[InspectorName] = pass1Inspector.Name
      ,InspectorSignature = InspectorSignature.Signature
      ,h.TestDate
from HeatTransferWash h
inner join SciProduction_Orders o on h.OrderID = o.ID
left join SciProduction_Style s on o.StyleUkey = s.Ukey
left join SciProduction_Pass1 p on h.EditName = p.ID
left join Pass1 p2 on h.EditName = p2.ID
left join SciProduction_Pass1 LastEdit1 on LastEdit1.ID = h.EditName
left join Pass1 LastEdit2 on  LastEdit2.ID = h.EditName
left join SciProduction_Pass1 pass1Approver WITH(NOLOCK) on h.Approver = pass1Approver.ID
left join SciProduction_Pass1 pass1Inspector WITH(NOLOCK) on h.Inspector = pass1Inspector.ID
outer apply (
    select top 1 Signature 
    from MainServer.Production.dbo.Technician t
    where Junk = 0 and ISNULL(p.ID, p2.ID) = t.ID
)Technician
outer apply (
    select top 1 Signature 
    from MainServer.Production.dbo.Technician t
    where Junk = 0 and h.Inspector = t.ID
)InspectorSignature
outer apply (
    select top 1 Signature 
    from MainServer.Production.dbo.Technician t
    where Junk = 0 and h.Approver = t.ID
)ApproverSignature
outer apply (
    select top 1 ID 
    from MainServer.Production.dbo.SubProcess t
    where Junk = 0 and t.ArtworkTypeID = h.ArtworkTypeID
)SubProcessInfo
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

            objParameter.Add("@ReportNo", DbType.String, ReportNo);
            IList<HeatTransferWash_Detail_Result> res = ExecuteList<HeatTransferWash_Detail_Result>(CommandType.Text, SbSql, objParameter);

            return res.Any() ? res.ToList() : new List<HeatTransferWash_Detail_Result>();
        }

        public HeatTransferWash_Detail_Result GetLastDetailData(string HTRefNo)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select top 1 *
from HeatTransferWash_Detail
where HTRefNo = @HTRefNo
order by EditDate desc 
";

            objParameter.Add("@HTRefNo", DbType.String, HTRefNo);
            IList<HeatTransferWash_Detail_Result> res = ExecuteList<HeatTransferWash_Detail_Result>(CommandType.Text, SbSql, objParameter);

            return res.Any() ? res.ToList().FirstOrDefault() : new HeatTransferWash_Detail_Result();
        }

        public int Insert_HeatTransferWash(HeatTransferWash_Result Req, string MDivision, string UserID, out string NewReportNo)
        {
            NewReportNo = GetID(MDivision + "HW", "HeatTransferWash", DateTime.Today, 2, "ReportNo");
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ReportNo", DbType.String, NewReportNo } ,
                { "@OrderID", DbType.String, Req.OrderID ?? ""} ,
                { "@BrandID", DbType.String, Req.BrandID } ,
                { "@SeasonID", DbType.String, Req.SeasonID } ,
                { "@StyleID", DbType.String, Req.StyleID } ,
                { "@Article", DbType.String, Req.Article } ,
                { "@Line", DbType.String,  Req.Line ?? "" } ,
                { "@Result", DbType.String,  Req.Result ?? "" } ,
                { "@Machine", DbType.String,  Req.Machine ?? "" } ,
                { "@Remark", DbType.String, Req.Remark ?? "" } ,
                { "@ArtworkTypeID", DbType.String, Req.ArtworkTypeID ?? "" } ,
                { "@Inspector", DbType.String, Req.Inspector ?? "" } ,
                { "@Approver", DbType.String, Req.Approver ?? "" } ,
                { "@TestDate", DbType.DateTime, Req.TestDate } ,
                { "@ReceivedDate", DbType.DateTime, Req.ReceivedDate } ,
                //{ "@Temperature", DbType.Int32, Req.Temperature } ,
                //{ "@Time", DbType.Int32, Req.Time } ,
                //{ "@Pressure", DbType.Decimal, Req.Pressure } ,
                //{ "@PeelOff", DbType.String, Req.PeelOff ?? "" } ,
                //{ "@Cycles", DbType.Int32, Req.Cycles } ,
                //{ "@TemperatureUnit", DbType.Int32, Req.TemperatureUnit } ,
                { "@AddName", DbType.String, UserID ?? "" } ,
                { "@Approver", DbType.String, Req.Approver ?? "" } ,
                { "@Preparer", DbType.String, Req.Preparer ?? "" } ,
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
           ,OrderID
           ,BrandID
           ,SeasonID
           ,StyleID
           ,Article
           ,Line
           ,Machine
           ,Remark
           ,ReceivedDate
           ,ArtworkTypeID
           ,Status
           ,Result
           ,Inspector
           ,Approver
           ,TestDate
           ,AddName
           ,AddDate
           ,Approver
           ,Preparer
)
VALUES  (     
            @ReportNo
           ,@OrderID
           ,@BrandID
           ,@SeasonID
           ,@StyleID
           ,@Article
           ,@Line
           ,@Machine
           ,@Remark
           ,@ReceivedDate
           ,@ArtworkTypeID
           ,'New'
           ,@Result
           ,@Inspector
           ,@Approver
           ,@TestDate
           ,@AddName
           ,GETDATE()
           ,@Approver
           ,@Preparer
)

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
                { "@EditName", DbType.String, Req.EditName ?? "" } ,
                { "@Temperature", DbType.Int32, Req.Temperature } ,
                { "@Time", DbType.Int32, Req.Time } ,
                { "@Pressure", DbType.Decimal, Req.Pressure } ,
                { "@PeelOff", DbType.String, Req.PeelOff ?? "" } ,
                { "@Cycles", DbType.Int32, Req.Cycles } ,
                { "@TemperatureUnit", DbType.Int32, Req.TemperatureUnit } ,
                { "@SecondTime", DbType.Int32, Req.SecondTime } ,
            };

            SbSql.Append($@"
INSERT INTO dbo.HeatTransferWash_Detail
           (ReportNo
           ,FabricRefNo
           ,HTRefNo
           ,Result
           ,Remark
           ,EditName
           ,EditDate
           ,Temperature
           ,Time
           ,Pressure
           ,PeelOff
           ,Cycles
           ,TemperatureUnit
           ,SecondTime)
VALUES      (@ReportNo
           ,@FabricRefNo
           ,@HTRefNo
           ,@Result
           ,@Remark
           ,@EditName
           ,GETDATE()
           ,@Temperature
           ,@Time
           ,@Pressure
           ,@PeelOff
           ,@Cycles
           ,@TemperatureUnit
           ,@SecondTime)
");

            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }

        public int Update_HeatTransferWash(HeatTransferWash_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@Line", DbType.String, Req.Main.Line ?? string.Empty);
            objParameter.Add("@Machine", DbType.String, Req.Main.Machine ?? string.Empty);
            objParameter.Add("@Result", DbType.String, Req.Main.Result ?? string.Empty);
            objParameter.Add("@Remark", DbType.String, Req.Main.Remark ?? string.Empty);
            objParameter.Add("@ArtworkTypeID", DbType.String, Req.Main.ArtworkTypeID ?? string.Empty);
            objParameter.Add("@ReceivedDate", DbType.Date, Req.Main.ReceivedDate);
            objParameter.Add("@ReportDate", DbType.Date, Req.Main.ReportDate);
            //objParameter.Add("@Temperature", DbType.Int32, Req.Main.Temperature);
            //objParameter.Add("@Time", DbType.Int32, Req.Main.Time);
            //objParameter.Add("@Pressure", DbType.Decimal, Req.Main.Pressure);
            //objParameter.Add("@PeelOff", DbType.String, Req.Main.PeelOff ?? string.Empty);
            //objParameter.Add("@Cycles", DbType.Int32, Req.Main.Cycles);
            //objParameter.Add("@TemperatureUnit", DbType.Int32, Req.Main.TemperatureUnit);
            objParameter.Add("@Editname", DbType.String, Req.Main.EditName ?? string.Empty);
            objParameter.Add("@ReportNo", DbType.String, Req.Main.ReportNo ?? string.Empty);
            objParameter.Add("@Preparer", DbType.String, Req.Main.Preparer ?? string.Empty);
            objParameter.Add("@Inspector", DbType.String, Req.Main.Inspector ?? string.Empty);
            objParameter.Add("@Approver", DbType.String, Req.Main.Approver ?? string.Empty);
            objParameter.Add("@TestDate", DbType.Date, Req.Main.TestDate);
			
            if (Req.Main.TestBeforePicture != null) { objParameter.Add("@TestBeforePicture", Req.Main.TestBeforePicture); }
            else { objParameter.Add("@TestBeforePicture", System.Data.SqlTypes.SqlBinary.Null); }
            if (Req.Main.TestAfterPicture != null) { objParameter.Add("@TestAfterPicture", Req.Main.TestAfterPicture); }
            else { objParameter.Add("@TestAfterPicture", System.Data.SqlTypes.SqlBinary.Null); }

            string head = $@"
SET XACT_ABORT ON

Update HeatTransferWash
Set Line = @Line
    ,Machine = @Machine
    ,Result = @Result
    ,Remark = @Remark
    ,ArtworkTypeID = @ArtworkTypeID
    ,EditDate = GETDATE()
    ,Editname = @Editname
    ,ReceivedDate = @ReceivedDate
    ,ReportDate = @ReportDate
    ,Preparer  =  @Preparer
    ,Inspector = @Inspector
    ,Approver = @Approver
    ,TestDate = @TestDate
where ReportNo = @ReportNo


if not exists (select 1 from SciPMSFile_HeatTransferWash where ReportNo = @ReportNo)
begin
    INSERT INTO SciPMSFile_HeatTransferWash (ReportNo,TestBeforePicture,TestAfterPicture)
    VALUES (@ReportNo,@TestBeforePicture,@TestAfterPicture)
end
else
begin
    UPDATE SciPMSFile_HeatTransferWash
    SET
        TestBeforePicture=@TestBeforePicture
        ,TestAfterPicture=@TestAfterPicture
    WHERE ReportNo = @ReportNo
end

";


            return ExecuteNonQuery(CommandType.Text, head, objParameter);
        }
        public int Delete_HeatTransferWash(string ReportNo)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ReportNo", DbType.String, ReportNo ?? string.Empty);

            string head = $@"
SET XACT_ABORT ON

DELETE FROM  HeatTransferWash
where ReportNo = @ReportNo

DELETE FROM  HeatTransferWash_Detail
where ReportNo = @ReportNo

DELETE FROM  SciPMSFile_HeatTransferWash
where ReportNo = @ReportNo
";

            return ExecuteNonQuery(CommandType.Text, head, objParameter);
        }
        public int Confirm_HeatTransferWash(HeatTransferWash_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@Editname", DbType.String, Req.Main.EditName ?? string.Empty);
            objParameter.Add("@ReportNo", DbType.String, Req.Main.ReportNo ?? string.Empty);
            objParameter.Add("@Status", DbType.String, Req.Main.Status ?? string.Empty);

            string head = $@"

Update b
Set  Status = @Status
    ,Result = IIF( (select COUNT(1) from HeatTransferWash_Detail a where a.ReportNo=b.ReportNo and a.Result='Fail') > 0 , 'Fail' ,'Pass')
    ,EditDate = GETDATE()
    ,Editname = @Editname
from HeatTransferWash b
where ReportNo = @ReportNo

";

            return ExecuteNonQuery(CommandType.Text, head, objParameter);
        }
        public void Update_HeatTransferWash_Detail(HeatTransferWash_ViewModel Req)
        {
            List<HeatTransferWash_Detail_Result> oldData = this.GetDetailData(Req.Main.ReportNo).ToList();

            List<HeatTransferWash_Detail_Result> needUpdateDetailList =
                PublicClass.CompareListValue<HeatTransferWash_Detail_Result>(
                    Req.Details,
                    oldData,
                    "Ukey",
                    "FabricRefNo,HTRefNo,Result,Remark,Temperature,Time,Pressure,PeelOff,Cycles,TemperatureUnit,SecondTime");

            string insert = $@"
INSERT INTO dbo.HeatTransferWash_Detail
           (ReportNo
           ,FabricRefNo
           ,HTRefNo
           ,Result
           ,Remark
           ,EditName
           ,EditDate
           ,Temperature
           ,Time
           ,Pressure
           ,PeelOff
           ,Cycles
           ,TemperatureUnit
           ,SecondTime)
     VALUES(
            @ReportNo
           ,@FabricRefNo
           ,@HTRefNo
           ,@Result
           ,@Remark
           ,@EditName
           ,GETDATE()
           ,@Temperature
           ,@Time
           ,@Pressure
           ,@PeelOff
           ,@Cycles
           ,@TemperatureUnit
           ,@SecondTime)

";
            string update = $@"
UPDATE dbo.HeatTransferWash_Detail
SET  FabricRefNo=@FabricRefNo
    ,HTRefNo = @HTRefNo
    ,Result = @Result
    ,Remark = @Remark
    ,EditName = @EditName
    ,EditDate = GETDATE()
    ,Temperature = @Temperature
    ,Time = @Time
    ,Pressure = @Pressure
    ,PeelOff = @PeelOff
    ,Cycles = @Cycles
    ,TemperatureUnit = @TemperatureUnit
    ,SecondTime = @SecondTime
WHERE UKey = @Ukey

";
            string delete = @"
delete HeatTransferWash_Detail where Ukey = @Ukey
";

            foreach (var detailItem in needUpdateDetailList)
            {
                SQLParameterCollection listDetailPar = new SQLParameterCollection();
                listDetailPar.Add("@FabricRefNo", DbType.String, detailItem.FabricRefNo ?? string.Empty);
                listDetailPar.Add("@HTRefNo", DbType.String, detailItem.HTRefNo ?? string.Empty);
                listDetailPar.Add("@Result", DbType.String, detailItem.Result);
                listDetailPar.Add("@Remark", DbType.String, detailItem.Remark ?? string.Empty);
                listDetailPar.Add("@EditName", DbType.String, detailItem.EditName ?? string.Empty);
                listDetailPar.Add("@Temperature", DbType.Int32, detailItem.Temperature);
                listDetailPar.Add("@Time", DbType.Int32, detailItem.Time);
                listDetailPar.Add("@Pressure", DbType.Decimal, detailItem.Pressure);
                listDetailPar.Add("@PeelOff", DbType.String, detailItem.PeelOff ?? "");
                listDetailPar.Add("@Cycles", DbType.Int32, detailItem.Cycles);
                listDetailPar.Add("@TemperatureUnit", DbType.Int32, detailItem.TemperatureUnit);
                listDetailPar.Add("@SecondTime", DbType.Int32, detailItem.SecondTime);
                switch (detailItem.StateType)
                {
                    case DatabaseObject.Public.CompareStateType.Add:
                        listDetailPar.Add("@ReportNo", DbType.String, detailItem.ReportNo);
                        ExecuteNonQuery(CommandType.Text, insert, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Edit:
                        listDetailPar.Add("@Ukey",  detailItem.Ukey);

                        ExecuteNonQuery(CommandType.Text, update, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.Delete:
                        listDetailPar.Add("@Ukey", detailItem.Ukey);
                        ExecuteNonQuery(CommandType.Text, delete, listDetailPar);
                        break;
                    case DatabaseObject.Public.CompareStateType.None:
                        break;
                    default:
                        break;
                }
            }
        }

        public IList<SelectListItem> GetArtworkTypeList(Orders orders)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select distinct Text = a.ArtworkTypeID ,Value = a.ArtworkTypeID
from Style_Artwork a
inner join Style s ON a.StyleUkey = s.Ukey
where s.BrandID = @BrandID and s.SeasonID = @SeasonID  and s.ID = @StyleID
";

            objParameter.Add("@BrandID", DbType.String, orders.BrandID);
            objParameter.Add("@SeasonID", DbType.String, orders.SeasonID);
            objParameter.Add("@StyleID", DbType.String, orders.StyleID);
            IList<SelectListItem> res = ExecuteList<SelectListItem>(CommandType.Text, SbSql, objParameter);

            return res.Any() ? res.ToList() : new List<SelectListItem>();
        }
        public List<string> GetArtworkTypeOri(Orders orders)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string SbSql = $@"
select distinct pgl.Annotation 
from Pattern_GL　pgl
inner join Pattern p on p.id = pgl.id and p.Version = pgl.Version
where StyleUkey IN (
	select Ukey
	from Style s 
    where s.BrandID = @BrandID and s.SeasonID = @SeasonID  and s.ID = @StyleID
)
AND pgl.Annotation <> ''
";

            objParameter.Add("@BrandID", DbType.String, orders.BrandID);
            objParameter.Add("@SeasonID", DbType.String, orders.SeasonID);
            objParameter.Add("@StyleID", DbType.String, orders.StyleID);
            DataTable dt = ExecuteDataTable(CommandType.Text, SbSql, objParameter);

            return dt == null || dt.Rows.Count == 0 ? new List<string>() : dt.AsEnumerable().Select(o => o["Annotation"].ToString()).ToList<string>();
        }

        public DataTable GetReportTechnician(HeatTransferWash_Request Req)
        {
            SQLParameterCollection paras = new SQLParameterCollection();
            paras.Add("@ReportNo", Req.ReportNo);

            string sqlCmd = $@"
select Technician = ISNULL(mp.Name,pp.Name)
	   ,TechnicianSignture = t.Signature
from HeatTransferWash a
left join Pass1 mp on mp.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
left join MainServer.Production.dbo.Pass1 pp on pp.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
left join MainServer.Production.dbo.Technician t on t.ID = IIF(a.EditName = '' ,a.AddName ,a.EditName)
where a.ReportNo = @ReportNo
;

";
            return ExecuteDataTable(CommandType.Text, sqlCmd, paras);
        }
    }
}
