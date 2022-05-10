using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel.SampleRFT;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace ManufacturingExecutionDataAccessLayer.Provider.MSSQL
{
    public class InspectionBySPProvider : SQLDAL
    {
        #region 底層連線
        public InspectionBySPProvider(string ConString) : base(ConString) { }
        public InspectionBySPProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        public IList<SampleRFTInspection> Get_SampleRFTInspection(InspectionBySP_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection para = new SQLParameterCollection();
            SbSql.Append($@"
select a.*, b.StyleID ,b.SeasonID ,b.BrandID
from SampleRFTInspection a
INNER JOIN MainServer.Production.dbo.Orders b on a.OrderID=b.ID
where 1=1

");
            if (Req.ID > 0)
            {
                SbSql.Append($@" AND a.ID = @ID ");
                para.Add("@ID", Req.ID);
            }

            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql.Append($@" AND a.OrderID = @OrderID ");
                para.Add("@OrderID", Req.OrderID);
            }


            return ExecuteList<SampleRFTInspection>(CommandType.Text, SbSql.ToString(), para);
        }
        public IList<InspectionBySP_SearchResult> Get_SearchResults(InspectionBySP_ViewModel Req)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection para = new SQLParameterCollection();
            SbSql.Append($@"

select    OrderID = o.ID
        , o.CustPONo
        , o.StyleID
        , o.SeasonID
        , o.BrandID
        , Article = Articles.Val
        , SampleStage = o.OrderTypeID
from SciProduction_Orders o
OUTER APPLY(
	SELECT Val = STUFF((
		select DISTINCT ',' + oq.Article
		from SciProduction_Order_Qty oq
		where oq.ID=o.ID
		FOR XML PATH ('')
		),1,1,'')
)Articles
where o.Category='S' AND o.Junk=0 
AND o.FtyGroup ='ESP'
AND NOT EXISTS(
	select 1
	from SampleRFTInspection
	where Status='Confirmed' AND Result='Pass' AND OrderID=o.ID
)

");
            if (!string.IsNullOrEmpty(Req.OrderID))
            {
                SbSql.Append($@" AND o.ID = @OrderID ");
                para.Add("@OrderID", Req.OrderID);
            }

            if (!string.IsNullOrEmpty(Req.CustPONo))
            {
                SbSql.Append($@" AND o.CustPONo = @CustPONo ");
                para.Add("@CustPONo", Req.CustPONo);
            }

            if (!string.IsNullOrEmpty(Req.StyleID))
            {
                SbSql.Append($@" AND o.StyleID = @StyleID  ");
                para.Add("@StyleID", Req.StyleID);
            }

            if (!string.IsNullOrEmpty(Req.SeasonID))
            {
                SbSql.Append($@" AND o.SeasonID = @SeasonID  ");
                para.Add("@SeasonID", Req.SeasonID);
            }


            return ExecuteList<InspectionBySP_SearchResult>(CommandType.Text, SbSql.ToString(), para);
        }

        public IList<Orders> GetOrders(string OrderID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection para = new SQLParameterCollection();
            SbSql.Append($@"
select o.CustPONO, o.StyleID, o.SeasonID, o.BrandID, o.Model , o.OrderTypeID ,o.Qty , FactoryID = o.FtyGroup
    ,Article = Articles.Val
from SciProduction_Orders  o
OUTER APPLY(
	SELECT Val = STUFF((
		select DISTINCT ',' + oq.Article
		from SciProduction_Order_Qty oq
		where oq.ID=o.ID
		FOR XML PATH ('')
		),1,1,'')
)Articles
where o.ID = @OrderID
");

            para.Add("@OrderID", OrderID);

            return ExecuteList<Orders>(CommandType.Text, SbSql.ToString(), para);
        }

        public IList<SelectSewing> Get_SewingLineList(string FactoryID, string OrderID)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection para = new SQLParameterCollection();
            SbSql.Append($@"
select   SewingLine = ID
		,Selected = Cast( IIF(EXISTS(
			select 1
			from MainServer.Production.dbo.SewingSchedule_Detail sd
			where sd.OrderID = @OrderID AND sd.SewingLineID = s.ID
		),1,0) as bit)
from MainServer.Production.dbo.SewingLine s
where FactoryID = @FactoryID
");
            para.Add("@FactoryID", FactoryID);
            para.Add("@OrderID", OrderID);

            return ExecuteList<SelectSewing>(CommandType.Text, SbSql.ToString(), para);
        }

        public IList<Select_QC_InCharge> Get_QC_InChargeList()
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection para = new SQLParameterCollection();
            SbSql.Append($@"
select *
from SampleRFTInspection_QCInCharge
where Junk = 0
");

            return ExecuteList<Select_QC_InCharge>(CommandType.Text, SbSql.ToString(), para);
        }

        public IList<AcceptableQualityLevels> GetAcceptableQualityLevelsForSetting(string OrderID)
        {
            SQLParameterCollection para = new SQLParameterCollection();
            string sqlGetData = $@"

DECLARE @OrderQty as int = (
	select Qty
	from MainServer.Production.dbo.Orders 
	where id = @OrderID
)

select	InspectionLevels ,
		LotSize_Start	 ,
		LotSize_End		 ,
		SampleSize		 ,
		Ukey			 ,
		Junk			 ,
		AQLType			 ,
		AcceptedQty
from MainServer.Production.dbo.AcceptableQualityLevels WITH(NOLOCK)
where AQLType in (1,1.5,2.5) and InspectionLevels = 1 and AcceptedQty is not null 
AND LotSize_Start <= @OrderQty AND @OrderQty <= LotSize_End
order by AQLType , InspectionLevels
";
            para.Add("@OrderID", OrderID);
            return ExecuteList<AcceptableQualityLevels>(CommandType.Text, sqlGetData, para);
        }

        public IList<SampleRFTInspection> SettingSampleRFTInspection(InspectionBySP_Setting Req)
        {
            SQLParameterCollection para = new SQLParameterCollection();
            para.Add("@ID", Req.ID);
            para.Add("@OrderID", Req.OrderID);
            para.Add("@SewingLineID", Req.SewingLineID);
            para.Add("@InspectionDate", Req.InspectionDate);
            para.Add("@InspectionStage", Req.InspectionStage ?? string.Empty);
            para.Add("@InspectionTimes", Req.InspectionTimes ?? string.Empty);
            para.Add("@QCInCharge", Req.QCInCharge);
            para.Add("@AcceptableQualityLevelsUkey", DbType.Int64, Req.AcceptableQualityLevelsUkey);
            para.Add("@SampleSize", Req.SampleSize ?? 0);
            para.Add("@AcceptQty", Req.AcceptQty ?? 0);
            para.Add("@AddName", Req.AddName ?? string.Empty);
            para.Add("@EditName", Req.EditName ?? string.Empty);


            string sql = string.Empty;
            if (Req.ID <= 0)
            {
                sql = $@"
DECLARE @FactoryID as varchar(10)
DECLARE @StyleUkey as bigint
DECLARE @OrderQty as int

select @FactoryID = FtyGroup , @StyleUkey = StyleUkey , @OrderQty= Qty
from MainServer.Production.dbo.Orders 
where ID = @OrderID

INSERT INTO dbo.SampleRFTInspection
            (OrderID           ,SewingLineID           ,FactoryID
            ,StyleUkey           ,InspectionDate           ,OrderQty
            ,Status           ,InspectionStep           ,InspectionStage
            ,InspectionTimes           ,QCInCharge           ,AcceptableQualityLevelsUkey
            ,SampleSize           ,AcceptQty           ,AddDate           ,AddName)
        VALUES
            (@OrderID           ,@SewingLineID           ,@FactoryID
            ,@StyleUkey           ,@InspectionDate           ,@OrderQty
            ,'New'           ,'Insp-Setting'           ,@InspectionStage
            ,@InspectionTimes           ,@QCInCharge           ,@AcceptableQualityLevelsUkey
            ,@SampleSize           ,@AcceptQty           ,GETDATE()           ,@AddName)

SELECT CAST( @@IDENTITY  as bigint) as ID


";
            }
            else
            {
                sql = $@"
UPDATE SampleRFTInspection
SET  EditDate = GETDATE(), EditName = @EditName
    ,SewingLineID = @SewingLineID
    ,InspectionDate=@InspectionDate
    ,InspectionStage=@InspectionStage
    ,InspectionTimes=@InspectionTimes
    ,QCInCharge=@QCInCharge
    ,AcceptableQualityLevelsUkey=@AcceptableQualityLevelsUkey
    ,SampleSize=@SampleSize
    ,AcceptQty=@AcceptQty
WHERE ID = @ID


SELECT ID
FROM SampleRFTInspection
WHERE ID=@ID
";
            }

            return ExecuteList<SampleRFTInspection>(CommandType.Text, sql, para);
        }

        public void UpdateSampleRFTInspectionnByStep(SampleRFTInspection inspection, string currentStep, string userID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string sqlUpdCmd = string.Empty;

            switch (currentStep)
            {
                case "Insp-Setting":
                    sqlUpdCmd += $@"
update SampleRFTInspection
 set    InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @ID
";
                    objParameter.Add("@ID", inspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", inspection.InspectionStep);
                    break;
                case "Insp-CheckList":
                    sqlUpdCmd += $@"
update SampleRFTInspection
 set       CheckFabricApproval = @CheckFabricApproval
          ,CheckMetalDetection = @CheckMetalDetection
          ,CheckSealingSampleApproval = @CheckSealingSampleApproval
          ,CheckColorShade = @CheckColorShade
          ,CheckHandfeel = @CheckHandfeel
          ,CheckAppearance = @CheckAppearance
          ,CheckPrintEmbDecorations = @CheckPrintEmbDecorations
          ,CheckFiberContent = @CheckFiberContent
          ,CheckCareInstructions = @CheckCareInstructions
          ,CheckDecorativeLabel = @CheckDecorativeLabel
          ,CheckAdicomLabel = @CheckAdicomLabel
          ,CheckCountryofOrigion = @CheckCountryofOrigion
          ,CheckSizeKey = @CheckSizeKey
          ,Check8FlagLabel = @Check8FlagLabel
          ,CheckAdditionalLabel = @CheckAdditionalLabel
          ,CheckShippingMark = @CheckShippingMark
          ,CheckPolytagMarketing = @CheckPolytagMarketing
          ,CheckColorSizeQty = @CheckColorSizeQty
          ,CheckHangtag = @CheckHangtag
          ,InspectionStep = @InspectionStep
          ,EditName= @userID
          ,EditDate= getdate()
where   ID = @ID
";

                    objParameter.Add("@ID", inspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", inspection.InspectionStep);

                    objParameter.Add("@CheckFabricApproval", inspection.CheckFabricApproval);
                    objParameter.Add("@CheckMetalDetection", inspection.CheckMetalDetection);
                    objParameter.Add("@CheckSealingSampleApproval", inspection.CheckSealingSampleApproval);
                    objParameter.Add("@CheckColorShade", inspection.CheckColorShade);
                    objParameter.Add("@CheckHandfeel", inspection.CheckHandfeel);
                    objParameter.Add("@CheckAppearance", inspection.CheckAppearance);
                    objParameter.Add("@CheckPrintEmbDecorations", inspection.CheckPrintEmbDecorations);
                    objParameter.Add("@CheckFiberContent", inspection.CheckFiberContent);
                    objParameter.Add("@CheckCareInstructions", inspection.CheckCareInstructions);
                    objParameter.Add("@CheckDecorativeLabel", inspection.CheckDecorativeLabel);
                    objParameter.Add("@CheckAdicomLabel", inspection.CheckAdicomLabel);
                    objParameter.Add("@CheckCountryofOrigion", inspection.CheckCountryofOrigion);
                    objParameter.Add("@CheckSizeKey", inspection.CheckSizeKey);
                    objParameter.Add("@Check8FlagLabel", inspection.Check8FlagLabel);
                    objParameter.Add("@CheckAdditionalLabel", inspection.CheckAdditionalLabel);
                    objParameter.Add("@CheckShippingMark", inspection.CheckShippingMark);
                    objParameter.Add("@CheckPolytagMarketing", inspection.CheckPolytagMarketing);
                    objParameter.Add("@CheckColorSizeQty", inspection.CheckColorSizeQty);
                    objParameter.Add("@CheckHangtag", inspection.CheckHangtag);

                    break;
                case "Insp-Measurement":
                    sqlUpdCmd += $@"
update SampleRFTInspection
 set    InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @ID
";
                    objParameter.Add("@ID", inspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", inspection.InspectionStep);
                    break;
                case "Insp-AddDefect":
                    sqlUpdCmd += $@"
update SampleRFTInspection
 set    InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @ID
";
                    objParameter.Add("@ID", inspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", inspection.InspectionStep);
                    break;
                case "Insp-BA":
                    sqlUpdCmd += $@"
update SampleRFTInspection
 set    InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @ID
";
                    objParameter.Add("@ID", inspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", inspection.InspectionStep);
                    break;
                case "Insp-DummyFit":
                    sqlUpdCmd += $@"
update SampleRFTInspection
 set    InspectionStep = @InspectionStep,
        EditName= @userID,
        EditDate= getdate()
where   ID = @ID
";
                    objParameter.Add("@ID", inspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", inspection.InspectionStep);
                    break;
                case "Insp-Others":
                    sqlUpdCmd += $@"
update SampleRFTInspection
 set    InspectionStep = @InspectionStep,
        SubmitDate = GETDATE(),
        EditName= @userID,
        EditDate= getdate()
where   ID = @ID
";
                    objParameter.Add("@ID", inspection.ID);
                    objParameter.Add("@userID", userID);
                    objParameter.Add("@InspectionStep", inspection.InspectionStep);

                    break;

                default:
                    break;
            }

            ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }

        public List<string> GetArticleList(string OrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@OrderID", OrderID);

            string sqlGetMoistureArticleList = @"
select distinct oqd.Article 
from MainServer.Production.dbo.Order_Qty oqd with (nolock)
where oqd.ID = @OrderID 
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetMoistureArticleList, listPar);

            if (dtResult.Rows.Count == 0)
            {
                return new List<string>();
            }
            else
            {
                return dtResult.AsEnumerable().Select(s => s["Article"].ToString()).ToList();
            }

        }

        public IList<ArticleSize> GetArticleSizeList(string OrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@OrderID", OrderID);

            string sqlGetMoistureArticleList = @"

select distinct oqd.Article, oqd.SizeCode 
from MainServer.Production.dbo.Order_QtyShip_Detail oqd with (nolock)
where oqd.ID = @OrderID 
";

            return ExecuteList<ArticleSize>(CommandType.Text, sqlGetMoistureArticleList, listPar);
        }

        public List<string> GetProductTypeList(string OrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@SP", OrderID);

            string sqlGetMoistureArticleList = @"

----避免沒有Order_Location資料，預先塞入
DECLARE CUR_SewingOutput_Detail CURSOR FOR 
     Select orderid = @SP

declare @SP2 varchar(13) 
OPEN CUR_SewingOutput_Detail   
FETCH NEXT FROM CUR_SewingOutput_Detail INTO @SP2 
WHILE @@FETCH_STATUS = 0 
BEGIN
  exec MainServer.Production.dbo.Ins_OrderLocation @SP2
FETCH NEXT FROM CUR_SewingOutput_Detail INTO @SP2
END
CLOSE CUR_SewingOutput_Detail
DEALLOCATE CUR_SewingOutput_Detail


select distinct Location 
from MainServer.Production.dbo.Order_Location with (nolock)
where OrderId = @SP
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlGetMoistureArticleList, listPar);

            if (dtResult.Rows.Count == 0)
            {
                return new List<string>();
            }
            else
            {
                return dtResult.AsEnumerable().Select(s => s["Location"].ToString()).ToList();
            }
        }
        public string GetSizeUnitByCustPONO(string OrderID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                 { "@OrderID", DbType.String, OrderID} ,
            };

            string sqlcmd = @"
select TOP 1 SizeUnit 
from MainServer.Production.dbo.Style with (nolock)
where ukey IN (
select styleUkey
    from MainServer.Production.dbo.orders WITH(NOLOCK)
    WHERE ID = @OrderID
)
";

            DataTable dtResult = ExecuteDataTableByServiceConn(CommandType.Text, sqlcmd, objParameter);

            if (dtResult.Rows.Count > 0)
            {
                return dtResult.Rows[0]["SizeUnit"].ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        public IList<Measurement> GetMeasurementsByPOID(string OrderID, string userID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, OrderID },
                { "@userID", DbType.String, userID },
            };


            string sqlcmd = @"
declare @SizeUnit varchar(8)
declare @StyleUkey bigint = (
    select  StyleUkey
    from    Production.dbo.Orders with (nolock)
    where   ID = @OrderID
)


exec CopyStyle_ToMeasurement @userID,@StyleUkey;

select  @SizeUnit = SizeUnit
from    Production.dbo.Style WITH(NOLOCK)
where   Ukey = @StyleUkey

SELECT  a.StyleUkey
        ,a.Tol1
        ,a.Tol2
        ,a.Description
        ,a.Code
        ,a.SizeCode
        ,a.SizeSpec
        ,a.Ukey
        ,a.AddDate
        ,a.AddName
        ,a.Junk
        ,a.SizeGroup
        ,a.MeasurementTranslateUkey
        , [IsPatternMeas] = case when a.Description like '%pattern measn%'  then convert(bit, 1)
			                when  isnull(a.Tol1, '0') = '0' and isnull(a.Tol2, '0') = '0' and isnull(a.SizeSpec, '0') = '0' then convert(bit, 1)
			                when  isnull(a.SizeCode,'') = '' then convert(bit, 1)
			                when  a.SizeSpec like '%[a-zA-Z]%' then convert(bit, 1)
			                when  (UPPER(@SizeUnit) = 'INCH' and  a.SizeSpec like '%.%') then convert(bit, 1)
			                when  (UPPER(@SizeUnit) = 'CM' and  a.SizeSpec like '%/%') then convert(bit, 1)
			                else  convert(bit, 0) end
FROM [ManufacturingExecution].[dbo].[Measurement] a with(nolock)
where a.junk=0 
and a.StyleUkey = @StyleUkey 
order by a.Code

";

            return ExecuteList<Measurement>(CommandType.Text, sqlcmd, objParameter, 80);
        }

        /// <summary>
        /// Measurement圖片下拉選單資料來源
        /// </summary>
        /// <param name="OrderID"></param>
        /// <returns></returns>
        public IList<SelectListItem> Get_MeasurementImageSource(string OrderID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, OrderID } ,
            };
            string sqlcmd = @"
select Text= Cast( ROW_NUMBER() OVER(ORDER BY ID) as varchar)
        ,Value = Cast( ID as varchar)
from PMSFile.dbo.RFT_Inspection_Measurement

where OrderID = @OrderID
";
            return ExecuteList<SelectListItem>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<RFT_Inspection_Measurement_Image> Get_MeasurementImageList(string OrderID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", DbType.String, OrderID } ,
            };
            string sqlcmd = @"
select *
	,Seq = ROW_NUMBER() OVER(ORDER BY ID DESC)
from PMSFile.dbo.RFT_Inspection_Measurement
where OrderID = @OrderID
ORDER BY ID DESC
";
            return ExecuteList<RFT_Inspection_Measurement_Image>(CommandType.Text, sqlcmd, objParameter);
        }


        public int InsertRFT_Inspection_Measurement(InspectionBySP_Measurement measurement)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            List<RFT_Inspection_Measurement> MeasurementList = measurement.ListRFTMeasurementItem;
            string strNo = string.Empty;

            DataTable dt = ExecuteDataTableByServiceConn(CommandType.Text, $@"
select no = isnull(max(no),0)+1 from  ManufacturingExecution.dbo.RFT_Inspection_Measurement  WITH (NOLOCK) where styleukey = {measurement.StyleUkey}", objParameter);
            if (dt != null || dt.Rows.Count > 0)
            {
                strNo = dt.Rows[0]["no"].ToString();
            }

            int rowSeq = 1;
            int imgIdx = 0;
            string sqlcmd = $@"
SET XACT_ABORT ON
declare @AddDate datetime = GetDate()
";
            foreach (var item in MeasurementList)
            {
                objParameter.Add($"@MeasurementUkey{rowSeq}", item.MeasurementUkey);
                objParameter.Add($"@StyleUkey{rowSeq}", measurement.StyleUkey);
                objParameter.Add($"@No", strNo);
                objParameter.Add($"@Code{rowSeq}", item.Code);
                objParameter.Add($"@SizeCode{rowSeq}", item.SizeCode);
                objParameter.Add($"@SizeSpec{rowSeq}", item.ResultSizeSpec);
                objParameter.Add($"@OrderID{rowSeq}", measurement.OrderID);
                objParameter.Add($"@OrderID", measurement.OrderID);
                objParameter.Add($"@Article{rowSeq}", item.Article);
                objParameter.Add($"@Location{rowSeq}", item.Location);
                objParameter.Add($"@Line{rowSeq}", measurement.SewingLineID);
                objParameter.Add($"@FactoryID{rowSeq}", measurement.FactoryID);


                // 若沒填入SizeSpec則不insert RFT_Inspection_Measurement
                if (!string.IsNullOrEmpty(item.ResultSizeSpec))
                {
                    sqlcmd += $@"
insert into RFT_Inspection_Measurement(MeasurementUkey,StyleUkey,No,Code,SizeCode,SizeSpec,OrderID,Article,Location,Line,FactoryID,AddDate)
values(@MeasurementUkey{rowSeq},@StyleUkey{rowSeq},@No,@Code{rowSeq},@SizeCode{rowSeq},@SizeSpec{rowSeq},@OrderID{rowSeq},@Article{rowSeq},@Location{rowSeq},@Line{rowSeq},@FactoryID{rowSeq},@AddDate)
";
                }


                /*
                // 若沒填入SizeSpec，依然可填入RFT_Inspection_Measurement
                if (item.ImageList != null && item.ImageList.Any())
                {
                    foreach (var img in item.ImageList)
                    {
                        if (img != null)
                        {
                            objParameter.Add($"@Image{imgIdx}", DbType.Binary, img);
                            sqlcmd += $@"
insert into PMSFile.dbo.RFT_Inspection_Measurement(OrderID,Image)
values(@OrderID{rowSeq},@Image{imgIdx})
";
                            imgIdx++;
                        }
                    }
                }
                */
                rowSeq++;

            }


            // 若沒填入SizeSpec，依然可填入RFT_Inspection_Measurement
            if (measurement.ImageList != null && measurement.ImageList.Any())
            {
                sqlcmd += $@"
DELETE FROM PMSFile.dbo.RFT_Inspection_Measurement
where OrderID = @OrderID";

                foreach (var img in measurement.ImageList)
                {
                    // 圖片全部刪除後再寫入
                    if (img != null)
                    {
                        objParameter.Add($"@Image{imgIdx}", DbType.Binary, img);
                        //objParameter.Add($"@OrderID", measurement.OrderID);
                        sqlcmd += $@"
insert into PMSFile.dbo.RFT_Inspection_Measurement(OrderID,Image)
values(@OrderID,@Image{imgIdx})
";
                        imgIdx++;
                    }
                }
            }

            return ExecuteNonQuery(CommandType.Text, sqlcmd, objParameter);
        }


        public IList<SampleRFTInspection_Summary> GetDefectDefaultBody(long SampleRFTInspectionID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@SampleRFTInspectionID",  SampleRFTInspectionID } ,
            };
            string sqlcmd = $@"
select *
into #SampleRFTInspection_Detail
from SampleRFTInspection_Detail
where SampleRFTInspectionUkey = @SampleRFTInspectionID

select [RowIndex]=ROW_NUMBER() OVER(ORDER BY gdt.id,gdc.id) -1
    , ID = {SampleRFTInspectionID}
    , UKey = (
		select top 1 Ukey
		from #SampleRFTInspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
	)
    , GarmentDefectTypeID=gdt.ID
	, DefectType = gdt.Description
    , DefectTypeDesc = gdt.ID +'-'+gdt.Description

	, GarmentDefectCodeID  = gdc.ID
	, DefectCode = gdc.Description
    , DefectCodeDesc = gdc.ID +'-'+gdc.Description

	,AreaCodes = ISNULL( Areas.Val ,'')
	,Qty= Qty.Val
	,Responsibility = (
		select top 1 Responsibility
		from #SampleRFTInspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
	)
from Production.dbo.GarmentDefectType gdt 
inner join Production.dbo.GarmentDefectCode gdc on gdt.id=gdc.GarmentDefectTypeID 
outer apply(

	select Val = stuff((
		select DISTINCT ',' +a.AreaCode
		from #SampleRFTInspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
		FOR XML PATH('')
		),1,1,'')
)Areas
outer apply(
	select Val=COUNT(1)
	from #SampleRFTInspection_Detail a
	where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
)Qty
where 1=1 and gdt.Junk =0 and gdc.Junk =0
order by gdt.id,gdc.id

drop table #SampleRFTInspection_Detail

";
            return ExecuteList<SampleRFTInspection_Summary>(CommandType.Text, sqlcmd, objParameter);
        }
        public IList<Area> GetArea()
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                // { "@SampleRFTInspectionID",  SampleRFTInspectionID } ,
            };
            string sqlcmd = @"
select *
from Area
where Junk = 0
ORDER BY Type,Seq
";
            return ExecuteList<Area>(CommandType.Text, sqlcmd, objParameter);
        }
        public IList<SelectListItem> Get_DefectImageSource(long ID, long SampleRFTInspectionDetailUKey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@SampleRFTInspectionUkey", ID } ,
                { "@SampleRFTInspectionDetailUKey", SampleRFTInspectionDetailUKey } ,
            };
            string whereDetailUKey = string.Empty;

            if (SampleRFTInspectionDetailUKey > 0)
            {
                whereDetailUKey = "AND a.UKey = @SampleRFTInspectionDetailUKey";
            }

            string sqlcmd = $@"
select *
into #SampleRFTInspection_Detail
from SampleRFTInspection_Detail
where SampleRFTInspectionUkey = @SampleRFTInspectionUkey

select [RowIndex]=ROW_NUMBER() OVER(ORDER BY gdt.id,gdc.id) -1
    , UKey = (
		select top 1 Ukey
		from #SampleRFTInspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
	)
    , GarmentDefectTypeID=gdt.ID
	, DefectType = gdt.Description
    , DefectTypeDesc = gdt.ID +'-'+gdt.Description

	, GarmentDefectCodeID  = gdc.ID
	, DefectCode = gdc.Description
    , DefectCodeDesc = gdc.ID +'-'+gdc.Description

	,AreaCodes = ISNULL( Areas.Val ,'')
	,Qty= Qty.Val
	,Responsibility = (
		select top 1 Responsibility
		from #SampleRFTInspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
	)
into #base
from Production.dbo.GarmentDefectType gdt 
inner join Production.dbo.GarmentDefectCode gdc on gdt.id=gdc.GarmentDefectTypeID 
outer apply(

	select Val = stuff((
		select DISTINCT ',' +a.AreaCode
		from #SampleRFTInspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
		FOR XML PATH('')
		),1,1,'')
)Areas
outer apply(
	select Val=COUNT(1)
	from #SampleRFTInspection_Detail a
	where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
)Qty
where 1=1 and gdt.Junk =0 and gdc.Junk =0
order by gdt.id,gdc.id

select Text= Cast( ROW_NUMBER() OVER(ORDER BY a.GarmentDefectCodeID) as varchar)
        ,Value = a.GarmentDefectCodeID

from #base a
inner join  PMSFile.dbo.SampleRFTInspection_Detail b on a.UKey=b.SampleRFTInspectionDetailUKey
where 1=1
{whereDetailUKey}

drop table #SampleRFTInspection_Detail,#base
";
            return ExecuteList<SelectListItem>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<DefectImage> Get_DefectImageList(long ID, long SampleRFTInspectionDetailUKey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@SampleRFTInspectionUkey", ID } ,
                { "@SampleRFTInspectionDetailUKey", SampleRFTInspectionDetailUKey } ,
            };
            string whereDetailUKey = string.Empty;

            if (SampleRFTInspectionDetailUKey > 0)
            {
                whereDetailUKey = "AND a.UKey = @SampleRFTInspectionDetailUKey";
            }

            string sqlcmd = $@"
select *
into #SampleRFTInspection_Detail
from SampleRFTInspection_Detail
where SampleRFTInspectionUkey = @SampleRFTInspectionUkey

select [RowIndex]=ROW_NUMBER() OVER(ORDER BY gdt.id,gdc.id) -1
    , UKey = (
		select top 1 Ukey
		from #SampleRFTInspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
	)
    , GarmentDefectTypeID=gdt.ID
	, DefectType = gdt.Description
    , DefectTypeDesc = gdt.ID +'-'+gdt.Description

	, GarmentDefectCodeID  = gdc.ID
	, DefectCode = gdc.Description
    , DefectCodeDesc = gdc.ID +'-'+gdc.Description

	,AreaCodes = ISNULL( Areas.Val ,'')
	,Qty= Qty.Val
	,Responsibility = (
		select top 1 Responsibility
		from #SampleRFTInspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
	)
into #base
from Production.dbo.GarmentDefectType gdt 
inner join Production.dbo.GarmentDefectCode gdc on gdt.id=gdc.GarmentDefectTypeID 
outer apply(

	select Val = stuff((
		select DISTINCT ',' +a.AreaCode
		from #SampleRFTInspection_Detail a
		where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
		FOR XML PATH('')
		),1,1,'')
)Areas
outer apply(
	select Val=COUNT(1)
	from #SampleRFTInspection_Detail a
	where a.GarmentDefectTypeID=gdt.ID AND a.GarmentDefectCodeID=gdc.ID
)Qty
where 1=1 and gdt.Junk =0 and gdc.Junk =0
order by gdt.id,gdc.id

select a.*
    ,Seq = ROW_NUMBER() OVER(ORDER BY a.Ukey DESC)
    , SampleRFTInspectionDetailUKey = a.Ukey
    , ImageUKey = b.Ukey
    , b.Image
from #base a
inner join  PMSFile.dbo.SampleRFTInspection_Detail b on a.UKey=b.SampleRFTInspectionDetailUKey
where 1=1
{whereDetailUKey}

drop table #SampleRFTInspection_Detail,#base


";
            return ExecuteList<DefectImage>(CommandType.Text, sqlcmd, objParameter);
        }

        /// <summary>
        /// 針對現有AddDefect表身的資料異動處理
        /// </summary>
        /// <param name="details"></param>
        public int AddDefect_Detail_Process(List<SampleRFTInspection_Detail> details)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            var group = details.Select(o => o.GarmentDefectCodeID).Distinct();
            string FinalSql = string.Empty;

            int idx = 0;
            foreach (var key in group)
            {
                var datas = details.Where(o => o.GarmentDefectCodeID == key).ToList();
                long ID = details.FirstOrDefault().SampleRFTInspectionUkey;
                string GarmentDefectCodeID = datas.FirstOrDefault().GarmentDefectCodeID;

                string tmpTable = string.Empty;

                int count = 1;
                foreach (SampleRFTInspection_Detail data in datas)
                {
                    string tmp = $@"
SELECT [SampleRFTInspectionUkey]={data.SampleRFTInspectionUkey}
    ,[Ukey]=IIF (EXISTS(select TOP 1 Ukey from SampleRFTInspection_Detail where SampleRFTInspectionUKey={data.SampleRFTInspectionUkey} AND GarmentDefectCodeID = '{data.GarmentDefectCodeID}'AND AreaCode='{data.AreaCode}')
				,(select TOP 1 Ukey from SampleRFTInspection_Detail where SampleRFTInspectionUKey={data.SampleRFTInspectionUkey} AND GarmentDefectCodeID = '{data.GarmentDefectCodeID}' AND AreaCode='{data.AreaCode}')
				,0)
    ,[DefectCode]='{data.DefectCode}'
    ,[AreaCode]='{data.AreaCode}'
    ,[GarmentDefectTypeID]='{data.GarmentDefectTypeID}'
    ,[GarmentDefectCodeID]='{data.GarmentDefectCodeID}'
    ,[Qty]={data.Qty}
    ,[Responsibility]='{data.Responsibility}'
    ,[AddDate]=GETDATE()
";

                    tmpTable += tmp + Environment.NewLine;

                    if (count == 1)
                    {
                        tmpTable += $"INTO #source{idx}" + Environment.NewLine;
                    }

                    if (datas.Count > count)
                    {
                        tmpTable += "UNION" + Environment.NewLine;
                    }

                    count++;
                }

                objParameter.Add($"@SampleRFTInspectionUkey{idx}", ID);
                objParameter.Add($"@GarmentDefectCodeID{idx}", GarmentDefectCodeID);                

                // 判斷是否數量相同
                string sql = $@"
{tmpTable}

----新增 / 修改
MERGE SampleRFTInspection_Detail t 
USING #source{idx} s
    ON t.UKey = s.UKey 
WHEN MATCHED THEN UPDATE SET
	   t.DefectCode = s.DefectCode
      ,t.AreaCode = s.AreaCode
      ,t.GarmentDefectTypeID = s.GarmentDefectTypeID
      ,t.GarmentDefectCodeID = s.GarmentDefectCodeID
      ,t.Qty = s.Qty
      ,t.Responsibility = s.Responsibility
      ,t.AddDate = s.AddDate
when not matched by target then
	INSERT (SampleRFTInspectionUkey,DefectCode,AreaCode,GarmentDefectTypeID,GarmentDefectCodeID,Qty,Responsibility,AddDate)
	VALUES (s.SampleRFTInspectionUkey ,s.DefectCode ,s.AreaCode ,s.GarmentDefectTypeID ,s.GarmentDefectCodeID ,s.Qty ,s.Responsibility ,s.AddDate)
;

----整理出即將刪除的detail
Select t.*
INTO #deleteTarget{idx}
FROM SampleRFTInspection_Detail t
where SampleRFTInspectionUkey = @SampleRFTInspectionUkey{idx} and  GarmentDefectCodeID = @GarmentDefectCodeID{idx}
AND NOT EXISTS(
	SELECT *
	FROM #source{idx} s
	where t.SampleRFTInspectionUkey=s.SampleRFTInspectionUkey AND t.GarmentDefectCodeID=s.GarmentDefectCodeID
	AND t.AreaCode = s.AreaCode
)


--避免圖片會找不到Detail，因此先把圖片的Detail UKey改成沒有被刪除的Detail Ukey
UPDATE t
SET t.SampleRFTInspectionDetailUKey = (
	select TOP 1 UKey 
	from SampleRFTInspection_Detail a
	where EXISTS(
		select * 
		from #deleteTarget{idx} b
		where a.SampleRFTInspectionUkey=b.SampleRFTInspectionUkey AND a.GarmentDefectCodeID=b.GarmentDefectCodeID
		AND a.Ukey != b.Ukey
	)
)
from PMSFile.dbo.SampleRFTInspection_Detail t
where EXISTS(
	select * 
	from #deleteTarget{idx} s
	where t.SampleRFTInspectionDetailUKey = s.Ukey
)

--正式刪除
DELETE t
from SampleRFTInspection_Detail t
INNER JOIN #deleteTarget{idx} s on t.Ukey=s.Ukey

DROP TABLE #source{idx}, #deleteTarget{idx}
";

                FinalSql += sql;
                idx++;
            }

            return ExecuteNonQuery(CommandType.Text, FinalSql, objParameter);
        }

        public int AddDefect_Delete_Image(List<DefectImage> defectImages)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string ukeys = string.Join(",", defectImages.Select(o => o.ImageUKey));



            string Sql = "SET XACT_ABORT ON";
            if (defectImages.Count != 0)
            {
                Sql += $@"
delete from PMSFile.dbo.SampleRFTInspection_Detail
where Ukey IN ({ukeys})
";
            }

            return ExecuteNonQuery(CommandType.Text, Sql, objParameter);
        }

        public int AddDefect_Add_Image(SampleRFTInspection_Summary Req)
        {
            List<DefectImage> Images = Req.Images == null ? new List<DefectImage>() : Req.Images;
            long ID = Req.ID;

            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add($"@SampleRFTInspectionUkey", ID);

            var group = Images.Select(o => o.GarmentDefectCodeID).Distinct();

            string FinalSql = "SET XACT_ABORT ON";

            int idx = 0;
            int imageidx = 0;
            foreach (var key in group)
            {
                string GarmentDefectCodeID = key;
                var datas = Images.Where(o => o.GarmentDefectCodeID == key).ToList();

                objParameter.Add($"@GarmentDefectCodeID{idx}", GarmentDefectCodeID);

                foreach (DefectImage item in datas)
                {
                    objParameter.Add($"@Image{imageidx}", item.Image == null ? System.Data.SqlTypes.SqlBinary.Null : item.Image);
                    string Sql = $@"

INSERT PMSFile.dbo.SampleRFTInspection_Detail 
    (SampleRFTInspectionDetailUKey, Image)
VALUES
    ( (
            select TOP 1 Ukey 
            from SampleRFTInspection_Detail WITH(NOLOCK)
            where SampleRFTInspectionUkey = @SampleRFTInspectionUkey AND GarmentDefectCodeID = @GarmentDefectCodeID{idx}
        )
    , @Image{imageidx}
    )
";

                    FinalSql += Sql;
                    imageidx++;
                }
                idx++;
            }

            return ExecuteNonQuery(CommandType.Text, FinalSql, objParameter);
        }

        public int AddDefect_Update_Head(SampleRFTInspection inspection)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection(); 

            string sqlUpdCmd = $@"
update SampleRFTInspection
set    RejectQty = @RejectQty
      ,PassQty = SampleSize - @RejectQty
where   ID = @ID
";
            objParameter.Add("@ID", inspection.ID);
            objParameter.Add("@RejectQty", inspection.RejectQty);

            return ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }


        public IList<BACriteriaItem> GetBeautifulProductAudit(long ID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();

            listPar.Add("@ID", ID);

            string sqlGetData = $@"
select ID, Description 
into #baseBACriteria
from  Production.dbo.DropDownList ddl  WITH(NOLOCK)
where Type = 'PMS_BACriteria'
order by Seq

select  [Ukey] = isnull(fn.Ukey, 0),
        [BACriteria] = bac.ID,
        [BACriteriaDesc] = bac.Description,
        [Qty] = isnull(fn.Qty, 0),		
		[RowIndex]=ROW_NUMBER() OVER(ORDER BY bac.ID) -1
		

    from #baseBACriteria bac with (nolock)
    left join   SampleRFTInspection_NonBACriteria fn WITH(NOLOCK) on    fn.SampleRFTInspectionUkey = @ID
        and fn.BACriteria = bac.ID

DROP TABLE #baseBACriteria
";
            return ExecuteList<BACriteriaItem>(CommandType.Text, sqlGetData, listPar);
        }
        public IList<SelectListItem> Get_BAImageSource(long ID, long BAUkey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@SampleRFTInspectionUkey", ID } ,
                { "@SampleRFTInspection_NonBACriteria_UKey", BAUkey } ,
            };
            string sqlcmd = $@"

select Text= Cast( ROW_NUMBER() OVER(ORDER BY pmsfile.UKey) as varchar)
        ,Value = Cast( pmsfile.Ukey as varchar)
from SampleRFTInspection_NonBACriteria a
INNER JOIN PMSFile.dbo.SampleRFTInspection_NonBACriteriaImage pmsfile ON a.SampleRFTInspectionUkey=pmsfile.SampleRFTInspectionUkey AND a.Ukey=pmsfile.NonBACriteriaUkey
where a.SampleRFTInspectionUkey = @SampleRFTInspectionUkey
AND a.Ukey = @SampleRFTInspection_NonBACriteria_UKey

";
            return ExecuteList<SelectListItem>(CommandType.Text, sqlcmd, objParameter);
        }
        public IList<BAImage> Get_BAImageList(long ID, long BAUkey)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@SampleRFTInspectionUkey", ID } ,
                { "@SampleRFTInspection_NonBACriteria_UKey", BAUkey } ,
            };

            string sqlcmd = $@"
select ImageUKey = pmsfile.Ukey
	, a.BACriteria
	, SampleRFTInspection_NonBACriteriaUKey = a.Ukey
	, Seq = ROW_NUMBER() OVER(ORDER BY pmsfile.Ukey )
	, pmsfile.Image
from SampleRFTInspection_NonBACriteria a
INNER JOIN PMSFile.dbo.SampleRFTInspection_NonBACriteriaImage pmsfile ON a.SampleRFTInspectionUkey=pmsfile.SampleRFTInspectionUkey AND a.Ukey=pmsfile.NonBACriteriaUkey
where a.SampleRFTInspectionUkey = @SampleRFTInspectionUkey
AND a.Ukey = @SampleRFTInspection_NonBACriteria_UKey
";
            return ExecuteList<BAImage>(CommandType.Text, sqlcmd, objParameter);
        }

        public int BA_Detail_Process(List<BACriteriaItem> details)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            var group = details.Select(o => o.BACriteria).Distinct().ToList();
            string FinalSql = string.Empty;

            foreach (var detail in details)
            {
                string sql = $@"";

                if (detail.Ukey > 0)
                {
                    sql = $@"
UPDATE dbo.SampleRFTInspection_NonBACriteria
   SET Qty = {detail.Qty}
 WHERE Ukey = {detail.Ukey}
";
                }
                else
                {
                    sql = $@"
INSERT INTO dbo.SampleRFTInspection_NonBACriteria
           (SampleRFTInspectionUkey
           ,BACriteria
           ,Qty)
     VALUES
           ({detail.ID}
           ,'{detail.BACriteria}'
           ,{detail.Qty} )
";
                }
                FinalSql += sql;
            }
            return ExecuteNonQuery(CommandType.Text, FinalSql, objParameter);
        }
        public int BA_Delete_Image(List<BAImage> baImages)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            string ukeys = string.Join(",", baImages.Select(o => o.ImageUKey));



            string Sql = "SET XACT_ABORT ON";
            if (baImages.Count != 0)
            {
                Sql += $@"
delete from PMSFile.dbo.SampleRFTInspection_NonBACriteriaImage
where Ukey IN ({ukeys})
";
            }

            return ExecuteNonQuery(CommandType.Text, Sql, objParameter);
        }
        public int BA_Add_Image(BACriteriaItem Req)
        {
            List<BAImage> Images = Req.Images == null ? new List<BAImage>() : Req.Images;
            long ID = Req.ID;

            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add($"@SampleRFTInspectionUkey", ID);

            var group = Images.Select(o => o.BACriteria).Distinct().ToList();

            string FinalSql = "SET XACT_ABORT ON";

            int idx = 0;
            int imageidx = 0;
            foreach (var key in group)
            {
                string BACriteria = key;
                var datas = Images.Where(o => o.BACriteria == key).ToList();

                objParameter.Add($"@BACriteria{idx}", BACriteria);

                foreach (BAImage item in datas)
                {
                    objParameter.Add($"@Image{imageidx}", item.Image == null ? System.Data.SqlTypes.SqlBinary.Null : item.Image);
                    string Sql = $@"

INSERT PMSFile.dbo.SampleRFTInspection_NonBACriteriaImage 
    (SampleRFTInspectionUkey
    ,NonBACriteriaUkey
    ,Image)
VALUES
    ( @SampleRFTInspectionUkey
        ,(
            select TOP 1 Ukey 
            from SampleRFTInspection_NonBACriteria WITH(NOLOCK)
            where SampleRFTInspectionUkey = @SampleRFTInspectionUkey AND BACriteria = @BACriteria{idx}
        )
    , @Image{imageidx}
    )
";

                    FinalSql += Sql;
                    imageidx++;
                }
                idx++;
            }

            return ExecuteNonQuery(CommandType.Text, FinalSql, objParameter);
        }


        public int BA_Update_Head(SampleRFTInspection inspection)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlUpdCmd = $@"
update SampleRFTInspection
set    BAQty = @BAQty
where   ID = @ID
";
            objParameter.Add("@ID", inspection.ID);
            objParameter.Add("@BAQty", inspection.BAQty);

            return ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }

        public IList<DummyFitImage> GetDummyFitImages(string OrderID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection
            {
                { "@OrderID", OrderID } ,
            };

            string sqlcmd = $@"
select *
from PMSFile.dbo.RFT_PicDuringDummyFitting  WITH(NOLOCK)
where OrderID = @OrderID
";

            return ExecuteList<DummyFitImage>(CommandType.Text, sqlcmd, objParameter);
        }

        public IList<CFTComments_Result> Get_CFT_OrderComments(string OrderID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@OrderID", DbType.String, OrderID);

            StringBuilder SbSql = new StringBuilder();
            SbSql.Append($@"
select  PMS_RFTCommentsID = d.ID
	,d.Description
	,aa.Comnments
	,d.Seq
from Production.dbo.DropdownList d
OUTER APPLY(
	select *
	from RFT_OrderComments r
	where  d.ID = r.PMS_RFTCommentsID AND r.OrderID = @OrderID
)aa
where d.Type='PMS_RFTComments'
ORDER BY d.Seq
");
            return ExecuteList<CFTComments_Result>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public IList<InspectionBySP_Others> Get_RFTInspection_Result(long ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@ID", ID);


            StringBuilder SbSql = new StringBuilder();
            SbSql.Append($@"
IF EXISTS(
    select 1
    from SampleRFTInspection_Detail
    where SampleRFTInspectionUkey = @ID
    and Qty>0
)
BEGIN
	SELECT InspectorResult = 'Fail'
END
ELSE
BEGIN
	SELECT InspectorResult = 'Pass'
END
");
            return ExecuteList<InspectionBySP_Others>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        public int OthersUpdate(InspectionBySP_Others Req)
        {
            List<CFTComments_Result> comments = Req.DataList;
            long ID = Req.ID;

            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@OrderID", Req.OrderID);

            string FinalSql = $@"";
            int idx = 0;
            foreach (var item in comments)
            {
                objParameter.Add($"@PMS_RFTCommentsID{idx}", item.PMS_RFTCommentsID);
                objParameter.Add($"@Comnments{idx}", item.Comnments);

                string sql = $@"
IF EXISTS(
    select 1
    from dbo.RFT_OrderComments
    where OrderID = @OrderID AND PMS_RFTCommentsID =@PMS_RFTCommentsID{idx}
)
begin
    update RFT_OrderComments
    set Comnments = @Comnments{idx}
    where OrderID = @OrderID AND PMS_RFTCommentsID = @PMS_RFTCommentsID{idx}
end
else
begin
    INSERT INTO RFT_OrderComments
               (OrderID           ,PMS_RFTCommentsID           ,Comnments)
    values
               (@OrderID           ,@PMS_RFTCommentsID{idx}          ,@Comnments{idx})

end

";
                FinalSql += sql;
                idx++;
            }



            return ExecuteNonQuery(CommandType.Text, FinalSql, objParameter);
        }
        public int Others_Update_Head(long ID)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();

            string sqlUpdCmd = $@"
update SampleRFTInspection
set    Result = IIF( EXISTS(   
                        select 1
                        from SampleRFTInspection_Detail
                        where SampleRFTInspectionUkey = @ID
                        and Qty>0)
                    ,'Fail' ,'Pass')
    ,Status = 'Confirmed'
where   ID = @ID
";
            objParameter.Add("@ID", ID);

            return ExecuteNonQuery(CommandType.Text, sqlUpdCmd, objParameter);
        }
    }
}
