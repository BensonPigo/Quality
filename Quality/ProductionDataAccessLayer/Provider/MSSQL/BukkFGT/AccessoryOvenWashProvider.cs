using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Data.SqlClient;
using System.Linq;
using ToolKit;
using System.Web.Mvc;
using DatabaseObject.ResultModel;
using DatabaseObject;
using System.Transactions;
using DatabaseObject.ViewModel.BulkFGT;

namespace ProductionDataAccessLayer.Provider.MSSQL.BukkFGT
{
    public class AccessoryOvenWashProvider : SQLDAL
    {
        #region 底層連線
        public AccessoryOvenWashProvider(string ConString) : base(ConString) { }
        public AccessoryOvenWashProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion


        public Accessory_ViewModel GetHead(string OrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", OrderID);

            string sqlCmd = @"
select OrderID= o.ID
	,o.StyleID
	,o.BrandID
	,o.SeasonID
	,EarliestCutDate = o.CutInLine
	,EarliestSCIDel = p.MinSciDelivery
	,TargetLeadTime =   CASE WHEN p.MinSciDelivery IS NULL THEN NULL
							 WHEN o.CutInline < (SELECT DATEADD(DAY, (SELECT MtlLeadTime from System WITH(NOLOCK)) ,'2021-10-30')) THEN  o.CutInline
							 ELSE  (SELECT DATEADD(DAY, (SELECT MtlLeadTime from System WITH(NOLOCK)) ,'2021-10-30'))
						END	
	,CompletionDate = IIF( p.AIRLabInspPercent = 100,CompletionDate.Val,NULL)
	,ArticlePercent = p.AIRLabInspPercent
	,MtlCmplt = p.Complete
	,Remark = p.AIRLaboratoryRemark
	,CreateBy = Concat (p.AddName, '-', c.Name, ' ', convert(varchar,  p.AddDate, 120) )
	,EditBy = Concat (p.EditName, '-', e.Name, ' ', convert(varchar,  p.EditDate, 120) )
from PO p WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = p.ID
left join Pass1 c WITH(NOLOCK) on p.AddName = c.ID
left join Pass1 e WITH(NOLOCK) on p.AddName = e.ID
OUTER APPLY(	
	SELECT Val = MAX(MaxDate) FROM (
		select MaxDate = MAX(OvenDate)
		from AIR_Laboratory WITH(NOLOCK)
		where POID= p.ID
		UNION
		select MaxDate =MAX(WashDate)
		from AIR_Laboratory WITH(NOLOCK)
		where POID= p.ID
	)a
)CompletionDate
where p.ID = @ID


";

            IList<Accessory_ViewModel> listResult = ExecuteList<Accessory_ViewModel>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult.Count == 0 ? new Accessory_ViewModel() : listResult.ToList()[0];
        }

        public IList<Accessory_Result> GetDetail(string OrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", OrderID);

            string sqlCmd = @"
select 
	 Seq = Concat (a.Seq1, ' ', a.Seq2)
	, al.ReportNo
	,r.ExportID
	,r.WhseArrival
	,a.SCIRefno
	,a.Refno
	,Supplier = Concat (a.SuppID, s.AbbEn)
	,ColorID = pc.SpecValue
	,SizeSpec = ps.SpecValue
	,a.ArriveQty
	,al.InspDeadline
	,OverAllResult = al.Result
	,al.NonOven
	,OvenResult = al.Oven
	,al.OvenScale
	,al.OvenDate
	,al.OvenInspector
	,al.OvenRemark
	,al.NonWash

    --701
	,WashResult = al.Wash
	,al.WashScale
	,al.WashDate
	,al.WashInspector
	,al.WashRemark

    --501
	,al.NonWashingFastness
	,WashingFastnessResult = al.WashingFastness
	,al.WashingFastnessReportDate
	,al.WashingFastnessInspector
	,al.WashingFastnessRemark

	,a.ReceivingID

	-----以下為藏在背後的Key值不會秀在畫面上-----
	,AIR_LaboratoryID = al.ID
	,a.Seq1
	,a.Seq2
from AIR_Laboratory al WITH(NOLOCK)
inner join Air a WITH(NOLOCK) ON a.ID = al.ID
left join Receiving r WITH(NOLOCK) on a.ReceivingID = r.ID
left join Supp s WITH(NOLOCK) ON s.ID = a.Suppid
left join Po_Supp_Detail psd WITH(NOLOCK) on a.POID = psd.ID and a.SEQ1=psd.SEQ1  and a.SEQ2=psd.SEQ2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join PO_Supp_Detail_Spec ps WITH(NOLOCK) on psd.POID = ps.ID and psd.SEQ1 = ps.SEQ1 and psd.SEQ2 = ps.SEQ2 and ps.SpecColumnID = 'Size'
where a.POID = @ID
ORDER BY a.SEQ1, a.SEQ2, r.ExportID

";
            return ExecuteList<Accessory_Result>(CommandType.Text, sqlCmd, listPar);
        }

        public int UpdateInspPercent(string POID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@POID", POID);

            string sqlCmd = $@" exec UpdateInspPercent 'AIRLab',@POID ";

            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }


        public int Update_AIR_Laboratory(Accessory_ViewModel Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", Req.OrderID);
            listPar.Add("@AIRLaboratoryRemark", Req.Remark ?? "");
            listPar.Add("@EditName", Req.EditBy);

            string sqlCmd = @"
UPDATE PO 
SET AIRLaboratoryRemark = @AIRLaboratoryRemark 
    ,EditDate=GETDATE() ,EditName=@EditName
WHERE ID = @ID
";
            int idx = 0;
            foreach (var data in Req.DataList)
            {
                sqlCmd += $@"
UPDATE AIR_Laboratory 
SET  NonOven = @NonOven_{idx} ,NonWash = @NonWash_{idx} ,NonWashingFastness = @NonWashingFastness_{idx}
,EditDate=GETDATE() ,EditName=@EditName
where POID = @ID
AND Seq1 = @Seq1_{idx}
AND Seq2 = @Seq2_{idx}
";

                listPar.Add($"@Seq1_{idx}", DbType.String, data.Seq1);
                listPar.Add($"@Seq2_{idx}", DbType.String, data.Seq2);
                listPar.Add($"@NonOven_{idx}",DbType.Boolean, data.NonOven);
                listPar.Add($"@NonWash_{idx}", DbType.Boolean, data.NonWash);
                listPar.Add($"@NonWashingFastness_{idx}", DbType.Boolean, data.NonWashingFastness);
                idx++;
            }


            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }

        public int Update_AIR_Laboratory_AllResult(Accessory_ViewModel Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", Req.OrderID);
            listPar.Add("@AIRLaboratoryRemark", Req.Remark ?? "");
            listPar.Add("@EditName", Req.EditBy);

            string sqlCmd = string.Empty;

            int idx = 0;
            foreach (var data in Req.DataList)
            {
                sqlCmd += $@"
UPDATE AIR_Laboratory 
SET  Result =   CASE    WHEN NonOven = 1 AND NonWash = 1 AND NonWashingFastness = 1 THEN 'Pass' --3個都不檢驗：直接整體PPass
	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWash = 1 AND NonWashingFastness = 1 THEN Oven                        --只有Oven檢驗
	                    WHEN NonWash = 0 AND WashEncode = 1 AND NonOven = 1 AND NonWashingFastness = 1 THEN Wash                        --只有Wash檢驗
	                    WHEN NonWashingFastness = 0 AND WashingFastnessEncode = 1 AND NonOven = 1 AND NonWash = 1 THEN WashingFastness  --只有WashingFastness檢驗

	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWash = 0 AND WashEncode = 1 THEN IIF( Oven = 'Pass' and Wash = 'Pass' , 'Pass' , 'Fail' )                                    --Oven + Wash要檢驗
	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWashingFastness= 0 AND WashingFastnessEncode = 1 THEN IIF( Oven = 'Pass' and WashingFastness = 'Pass' , 'Pass' , 'Fail' )    --Oven + WashingFastness要檢驗
	                    WHEN NonWash = 0 AND WashEncode = 1 AND NonWashingFastness= 0 AND WashingFastnessEncode = 1 THEN IIF( Wash = 'Pass' and WashingFastness = 'Pass' , 'Pass' , 'Fail' )    --Wash + WashingFastness要檢驗

	                    ELSE (	--3個都要檢驗，則套用以下判斷
			                    CASE WHEN OvenEncode != 1 OR WashEncode != 1 OR WashingFastnessEncode != 1 OR Oven='' OR Wash= '' OR WashingFastness = '' THEN ''   --其中一個未Encode或檢驗：  檢驗未完成，最終結果為空白
					                 WHEN Oven='Fail' OR Wash= 'Fail' OR WashingFastness = 'Fail'  THEN 'Fail'  --其中一個Fail：最終結果 Fail
			                    ELSE 'Pass'
			                    END
	                    )
                END
where POID = @ID
AND Seq1 = @Seq1_{idx}
AND Seq2 = @Seq2_{idx}
";

                listPar.Add($"@Seq1_{idx}", DbType.String, data.Seq1);
                listPar.Add($"@Seq2_{idx}", DbType.String, data.Seq2);
                listPar.Add($"@NonOven_{idx}", DbType.Boolean, data.NonOven);
                listPar.Add($"@NonWash_{idx}", DbType.Boolean, data.NonWash);
                listPar.Add($"@NonWashingFastness_{idx}", DbType.Boolean, data.NonWashingFastness);
                idx++;
            }

            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }

        #region Oven
        public Accessory_Oven GetOvenTest(Accessory_Oven Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = @"
select   al.POID
		,al.ReportNo
        ,a.SCIRefno
        ,WKNo = r.ExportId
        ,a.Refno
        ,a.ArriveQty
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,Unit = psd.StockUnit
        ,Color = pc.SpecValue
        ,Size = ps.SpecValue
        ,Scale = al.OvenScale
        ,OvenResult = al.Oven
        ,Remark = al.OvenRemark
        ,al.OvenInspector
        ,OvenInspectorName = q.Name
        ,al.OvenDate
        ,al.OvenEncode 
	
        ,OverAllResult = al.Result
        ,AIR_LaboratoryID = al.ID
        ,al.Seq1
        ,al.Seq2
		,ali.OvenTestBeforePicture
		,ali.OvenTestAfterPicture
from AIR_Laboratory al WITH(NOLOCK)
left join SciPMSFile_AIR_Laboratory ali WITH(NOLOCK) ON ali.ID=al.ID AND  ali.POID = al.POID AND ali.Seq1 = al.Seq1 AND ali.Seq2 = al.Seq2
inner join AIR a WITH(NOLOCK) ON a.ID = al.ID
left join Receiving r WITH(NOLOCK) on a.ReceivingID = r.Id
left join Supp s WITH(NOLOCK) on a.Suppid = s.ID
left join PO_Supp_Detail psd WITH(NOLOCK) ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join PO_Supp_Detail_Spec ps WITH(NOLOCK) on psd.POID = ps.ID and psd.SEQ1 = ps.SEQ1 and psd.SEQ2 = ps.SEQ2 and ps.SpecColumnID = 'Size'
left join Pass1 q WITH(NOLOCK) on q.ID = al.OvenInspector
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";
            IList<Accessory_Oven> listResult = ExecuteList<Accessory_Oven>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult.FirstOrDefault();
        }

        /// <summary>
        /// 取得匯出報表資訊
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public Accessory_OvenExcel GetOvenTestExcel(Accessory_Oven Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = @"
select   al.POID
		,o.BrandID
		,o.SeasonID
		,o.StyleID
        ,WKNo = r.ExportId
        ,a.Refno
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,Color = pc.SpecValue
        ,Size = ps.SpecValue
        ,OvenResult = al.Oven
        ,Remark = al.OvenRemark
        ,al.OvenInspector
        ,OvenInspectorName = q.Name
        ,al.OvenDate
        ,al.OvenEncode 
	    ,Seq = al.Seq1 + ' ' + al.Seq2
        ,OverAllResult = al.Result
        ,AIR_LaboratoryID = al.ID
        ,al.Seq1
        ,al.Seq2
		,ali.OvenTestBeforePicture
		,ali.OvenTestAfterPicture
        ,al.ReportNo
from AIR_Laboratory al WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = al.POID
left join SciPMSFile_AIR_Laboratory ali WITH(NOLOCK) ON ali.ID=al.ID AND  ali.POID = al.POID AND ali.Seq1 = al.Seq1 AND ali.Seq2 = al.Seq2
inner join AIR a WITH(NOLOCK) ON a.ID = al.ID
left join Receiving r WITH(NOLOCK) on a.ReceivingID = r.Id
left join Supp s WITH(NOLOCK) on a.Suppid = s.ID
left join PO_Supp_Detail psd WITH(NOLOCK) ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join PO_Supp_Detail_Spec ps WITH(NOLOCK) on psd.POID = ps.ID and psd.SEQ1 = ps.SEQ1 and psd.SEQ2 = ps.SEQ2 and ps.SpecColumnID = 'Size'
left join Pass1 q WITH(NOLOCK) on q.ID = al.OvenInspector
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";

            IList<Accessory_OvenExcel> listResult = ExecuteList<Accessory_OvenExcel>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult.FirstOrDefault();
        }

        public int UpdateOvenTest(Accessory_Oven Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string NewReportNo = GetID(Req.MDivisionID + "AT", "AIR_Laboratory", DateTime.Today, 2, "ReportNo");
            listPar.Add("@ReportNo", NewReportNo);

            string updateCol = string.Empty;
            string updatePicCol = string.Empty;
            #region 需要UPDATE的欄位
            if (!string.IsNullOrEmpty(Req.Scale))
            {
                updateCol += $@" , OvenScale = @Scale" + Environment.NewLine;
                listPar.Add("@Scale", Req.Scale);
            }


            updateCol += $@" , Oven = @OvenResult" + Environment.NewLine;
            listPar.Add("@OvenResult", Req.OvenResult ?? string.Empty);

            updateCol += $@" , OvenRemark = @Remark" + Environment.NewLine;
            listPar.Add("@Remark", Req.Remark ?? string.Empty);

            if (!string.IsNullOrEmpty(Req.OvenInspector))
            {
                updateCol += $@" , OvenInspector = @OvenInspector" + Environment.NewLine;
                listPar.Add("@OvenInspector", Req.OvenInspector);
            }

            updateCol += $@" , OvenDate = @OvenDate" + Environment.NewLine;
            if (Req.OvenDate.HasValue)
            {
                listPar.Add("@OvenDate", Req.OvenDate.Value);
            }
            else
            {
                listPar.Add("@OvenDate", DBNull.Value);
            }

            if (!string.IsNullOrEmpty(Req.EditName))
            {
                updateCol += $@" , EditName = @EditName" + Environment.NewLine;
                listPar.Add("@EditName", Req.EditName);
            }

            updatePicCol += $@"        ,OvenTestBeforePicture = @OvenTestBeforePicture " + Environment.NewLine;
            updatePicCol += $@"        ,OvenTestAfterPicture = @OvenTestAfterPicture " + Environment.NewLine;
            if (Req.OvenTestBeforePicture != null)
            {
                listPar.Add("@OvenTestBeforePicture", Req.OvenTestBeforePicture);
            }
            else
            {
                listPar.Add("@OvenTestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.OvenTestAfterPicture != null)
            {
                listPar.Add("@OvenTestAfterPicture", Req.OvenTestAfterPicture);
            }
            else
            {
                listPar.Add("@OvenTestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }
            #endregion

            string sqlCmd = $@"
SET XACT_ABORT ON

UPDATE AIR_Laboratory
SET EditDate=GETDATE()
{updateCol}
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2

UPDATE AIR_Laboratory
SET ReportNo = @ReportNo
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
    and ReportNo = ''

if not exists (select 1 from SciPMSFile_AIR_Laboratory where ID = @AIR_LaboratoryID and POID = @POID and Seq1 = @Seq1 and Seq2 = @Seq2)
begin
    INSERT INTO SciPMSFile_AIR_Laboratory (ID,POID,Seq1,Seq2,OvenTestBeforePicture,OvenTestAfterPicture)
    VALUES (@AIR_LaboratoryID,@POID,@Seq1,@Seq2,@OvenTestBeforePicture,@OvenTestAfterPicture)
end
else
begin
    UPDATE SciPMSFile_AIR_Laboratory
    SET POID=POID
    {updatePicCol}
    where   ID = @AIR_LaboratoryID
        and POID = @POID
        and Seq1 = @Seq1
        and Seq2 = @Seq2
end
";

            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }


        public Accessory_Oven Oven_EncodeCheck(Accessory_Oven Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = @"
select   1
from AIR_Laboratory
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
    AND Oven = ''
";
            IList<Accessory_Oven> listResult = ExecuteList<Accessory_Oven>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count > 0)
            {
                throw new Exception("Result cannot be empty.");
            }

            return listResult.FirstOrDefault();
        }

        public int EncodeAmendOven(Accessory_Oven Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string updateCol = string.Empty;
            #region 需要UPDATE的欄位


            updateCol += $@" , OvenEncode = @OvenEncode" + Environment.NewLine;
            listPar.Add("@OvenEncode", Req.OvenEncode);

            if (!string.IsNullOrEmpty(Req.EditName))
            {
                updateCol += $@" , EditName = @EditName" + Environment.NewLine;
                listPar.Add("@EditName", Req.EditName);
            }
            #endregion

            string sqlCmd = $@"
UPDATE AIR_Laboratory
SET EditDate=GETDATE()
{updateCol}
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
";


            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }

        /// <summary>
        /// 取得寄信資訊
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public DataTable GetData_OvenDataTable(Accessory_Oven Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            objParameter.Add("@POID", Req.POID);
            objParameter.Add("@Seq1", Req.Seq1);
            objParameter.Add("@Seq2", Req.Seq2);

            string sqlGetData = @"
select  [SP#] = al.POID
		,Style = o.StyleID
		,Brand = o.BrandID
		,Season = o.SeasonID
		,Seq = al.Seq1 + ' ' + al.Seq2
        ,[WK#] = r.ExportId
		,[Arrive W/H Date]= convert(varchar, r.WhseArrival , 111) 
        ,a.SCIRefno
        ,a.Refno
        ,Color = pc.SpecValue
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,a.ArriveQty
		,[Oven Result]=al.Oven
        ,[Oven Scale] = al.OvenScale
        ,[Oven Last Test Date]= convert(varchar, al.OvenDate , 111)  
		,[Oven Lab Tech	AIR_Laboratory]=al.OvenInspector
        ,Remark = al.OvenRemark
        -- ,[TestBeforePicture] = ali.OvenTestBeforePicture
        -- ,[TestAfterPicture] = ali.OvenTestAfterPicture
from AIR_Laboratory al WITH(NOLOCK)
inner join AIR a WITH(NOLOCK) ON a.ID = al.ID
INNER JOIn Orders o WITH(NOLOCK) ON o.ID = a.POID
-- left join SciPMSFile_AIR_Laboratory ali WITH(NOLOCK) ON ali.ID=al.ID AND  ali.POID = al.POID AND ali.Seq1 = al.Seq1 AND ali.Seq2 = al.Seq2
left join Receiving r WITH(NOLOCK) on a.ReceivingID = r.Id
left join Supp s WITH(NOLOCK) on a.Suppid = s.ID
left join PO_Supp_Detail psd WITH(NOLOCK) ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);
        }

        #endregion

        #region Wash
        public Accessory_Wash GetWashTest(Accessory_Wash Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = @"
select   al.POID
		,al.ReportNo
        ,a.SCIRefno
        ,WKNo = r.ExportId
        ,a.Refno
        ,a.ArriveQty
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,Unit = psd.StockUnit
        ,Color = pc.SpecValue
        ,Size = ps.SpecValue
        ,Scale = al.WashScale
        ,WashResult = al.Wash
        ,Remark = al.WashRemark
        ,al.WashInspector
        ,WashInspectorName = q.Name
        ,al.WashDate
        ,al.WashEncode 
	
        ,OverAllResult = al.Result
        ,AIR_LaboratoryID = al.ID
        ,al.Seq1
        ,al.Seq2
		,ali.WashTestBeforePicture
		,ali.WashTestAfterPicture
        ,al.MachineWash
        ,al.WashingTemperature
        ,al.DryProcess
        ,al.MachineModel
        ,al.WashingCycle

from AIR_Laboratory al WITH(NOLOCK)
left join SciPMSFile_AIR_Laboratory ali WITH(NOLOCK) ON ali.ID=al.ID AND  ali.POID = al.POID AND ali.Seq1 = al.Seq1 AND ali.Seq2 = al.Seq2
inner join AIR a WITH(NOLOCK) ON a.ID = al.ID
left join Receiving r WITH(NOLOCK) on a.ReceivingID = r.Id
left join Supp s WITH(NOLOCK) on a.Suppid = s.ID
left join PO_Supp_Detail psd WITH(NOLOCK) ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join PO_Supp_Detail_Spec ps WITH(NOLOCK) on psd.POID = ps.ID and psd.SEQ1 = ps.SEQ1 and psd.SEQ2 = ps.SEQ2 and ps.SpecColumnID = 'Size'
left join Pass1 q WITH(NOLOCK) on q.ID = al.WashInspector
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";
            IList<Accessory_Wash> listResult = ExecuteList<Accessory_Wash>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult.FirstOrDefault();
        }

        public int UpdateWashTest(Accessory_Wash Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string NewReportNo = GetID(Req.MDivisionID + "AT", "AIR_Laboratory", DateTime.Today, 2, "ReportNo");
            listPar.Add("@ReportNo", NewReportNo);

            string updateCol = string.Empty;
            string updatePicCol = string.Empty;
            #region 需要UPDATE的欄位
            if (!string.IsNullOrEmpty(Req.Scale))
            {
                updateCol += $@" , WashScale = @Scale" + Environment.NewLine;
                listPar.Add("@Scale", Req.Scale);
            }


            updateCol += $@" , Wash = @WashResult" + Environment.NewLine;
            listPar.Add("@WashResult", Req.WashResult ?? string.Empty);

            updateCol += $@" , WashRemark = @Remark" + Environment.NewLine;
            listPar.Add("@Remark", Req.Remark ?? string.Empty);

            if (!string.IsNullOrEmpty(Req.WashInspector))
            {
                updateCol += $@" , WashInspector = @WashInspector" + Environment.NewLine;
                listPar.Add("@WashInspector", Req.WashInspector);
            }

            updateCol += $@" , MachineWash = @MachineWash" + Environment.NewLine;
            listPar.Add("@MachineWash", Req.MachineWash ?? string.Empty);

            updateCol += $@" , WashingTemperature = @WashingTemperature" + Environment.NewLine;
            listPar.Add("@WashingTemperature", DbType.Int32, Req.WashingTemperature);

            updateCol += $@" , DryProcess = @DryProcess" + Environment.NewLine;
            listPar.Add("@DryProcess", DbType.String, Req.DryProcess ?? string.Empty);

            updateCol += $@" , MachineModel = @MachineModel" + Environment.NewLine;
            listPar.Add("@MachineModel", Req.MachineModel ?? string.Empty);

            updateCol += $@" , WashingCycle = @WashingCycle" + Environment.NewLine;
            listPar.Add("@WashingCycle", DbType.Int32, Req.WashingCycle);

            updateCol += $@" , WashDate = @WashDate" + Environment.NewLine;
            if (Req.WashDate.HasValue)
            {
                listPar.Add("@WashDate", Req.WashDate.Value);
            }
            else
            {
                listPar.Add("@WashDate", DBNull.Value);
            }

            if (!string.IsNullOrEmpty(Req.EditName))
            {
                updateCol += $@" , EditName = @EditName" + Environment.NewLine;
                listPar.Add("@EditName", Req.EditName);
            }

            updatePicCol += $@"        ,WashTestBeforePicture = @WashTestBeforePicture " + Environment.NewLine;
            updatePicCol += $@"        ,WashTestAfterPicture = @WashTestAfterPicture " + Environment.NewLine;
            if (Req.WashTestBeforePicture != null)
            {
                listPar.Add("@WashTestBeforePicture", Req.WashTestBeforePicture);
            }
            else
            {
                listPar.Add("@WashTestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.WashTestAfterPicture != null)
            {
                listPar.Add("@WashTestAfterPicture", Req.WashTestAfterPicture);
            }
            else
            {
                listPar.Add("@WashTestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }
            #endregion

            string sqlCmd = $@"
SET XACT_ABORT ON

UPDATE AIR_Laboratory
SET EditDate=GETDATE()
{updateCol}
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2

UPDATE AIR_Laboratory
SET ReportNo = @ReportNo
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
    and ReportNo=''

if not exists (select 1 from SciPMSFile_AIR_Laboratory where ID = @AIR_LaboratoryID and POID = @POID and Seq1 = @Seq1 and Seq2 = @Seq2)
begin
    INSERT INTO SciPMSFile_AIR_Laboratory (ID,POID,Seq1,Seq2,WashTestBeforePicture,WashTestAfterPicture)
    VALUES (@AIR_LaboratoryID,@POID,@Seq1,@Seq2,@WashTestBeforePicture,@WashTestAfterPicture)
end
else
begin
    UPDATE SciPMSFile_AIR_Laboratory
    SET POID=POID
    {updatePicCol}
    where   ID = @AIR_LaboratoryID
        and POID = @POID
        and Seq1 = @Seq1
        and Seq2 = @Seq2
end
";
            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }


        public Accessory_Wash Wash_EncodeCheck(Accessory_Wash Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = @"
select   1
from AIR_Laboratory
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
    AND Wash = ''
";
            IList<Accessory_Wash> listResult = ExecuteList<Accessory_Wash>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count > 0)
            {
                throw new Exception("Result cannot be empty.");
            }

            return listResult.FirstOrDefault();
        }
        public int EncodeAmendWash(Accessory_Wash Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string updateCol = string.Empty;
            #region 需要UPDATE的欄位


            updateCol += $@" , WashEncode = @WashEncode" + Environment.NewLine;
            listPar.Add("@WashEncode", Req.WashEncode);

            if (!string.IsNullOrEmpty(Req.EditName))
            {
                updateCol += $@" , EditName = @EditName" + Environment.NewLine;
                listPar.Add("@EditName", Req.EditName);
            }
            #endregion

            string sqlCmd = $@"
UPDATE AIR_Laboratory
SET EditDate=GETDATE()
{updateCol}
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
";


            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }

        /// <summary>
        /// 取得寄信資訊
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public DataTable GetData_WashDataTable(Accessory_Wash Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            objParameter.Add("@POID", Req.POID);
            objParameter.Add("@Seq1", Req.Seq1);
            objParameter.Add("@Seq2", Req.Seq2);

            string sqlGetData = @"
select  [SP#] = al.POID
		,Style = o.StyleID
		,Brand = o.BrandID
		,Season = o.SeasonID
		,Seq = al.Seq1 + ' ' + al.Seq2
        ,[WK#] = r.ExportId
		,[Arrive W/H Date]=  convert(varchar, r.WhseArrival , 111) 
        ,a.SCIRefno
        ,a.Refno
        ,Color = pc.SpecValue
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,a.ArriveQty
		,[Wash Result]=al.Wash
        ,[Wash Scale] = al.WashScale
        ,[Wash Last Test Date]=  convert(varchar, al.WashDate , 111)  
		,[Wash Lab Tech	AIR_Laboratory]=al.WashInspector
        ,Remark = al.WashRemark
	    --- , [TestBeforePicture] = ali.WashTestBeforePicture
	    -- , [TestAfterPicture] = ali.WashTestAfterPicture
from AIR_Laboratory al WITH(NOLOCK)
inner join AIR a WITH(NOLOCK) ON a.ID = al.ID
INNER JOIn Orders o WITH(NOLOCK) ON o.ID = a.POID
left join SciPMSFile_AIR_Laboratory ali WITH(NOLOCK) ON ali.ID=al.ID AND ali.POID = al.POID AND ali.Seq1 = al.Seq1 AND ali.Seq2 = al.Seq2
left join Receiving r WITH(NOLOCK) on a.ReceivingID = r.Id
left join Supp s WITH(NOLOCK) on a.Suppid = s.ID
left join PO_Supp_Detail psd WITH(NOLOCK) ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);
        }


        /// <summary>
        /// 取得匯出報表資訊
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public Accessory_WashExcel GetWashTestExcel(Accessory_Wash Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = @"
select   al.POID
		,o.BrandID
		,o.SeasonID
		,o.StyleID
        ,WKNo = r.ExportId
        ,a.Refno
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,Color = pc.SpecValue
        ,Size = ps.SpecValue
        ,WashResult = al.Wash
        ,Remark = al.WashRemark
        ,al.WashInspector
        ,WashInspectorName = q.Name
        ,al.WashDate
        ,al.WashEncode 
	    ,Seq = al.Seq1 + ' ' + al.Seq2
        ,OverAllResult = al.Result
        ,AIR_LaboratoryID = al.ID
        ,al.Seq1
        ,al.Seq2
		,ali.WashTestBeforePicture
		,ali.WashTestAfterPicture
        ,al.MachineWash
        ,al.WashingTemperature
        ,al.DryProcess
        ,al.MachineModel
        ,al.WashingCycle
        ,al.ReportNo
from AIR_Laboratory al WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = al.POID
left join SciPMSFile_AIR_Laboratory ali WITH(NOLOCK) ON ali.ID=al.ID AND  ali.POID = al.POID AND ali.Seq1 = al.Seq1 AND ali.Seq2 = al.Seq2
inner join AIR a WITH(NOLOCK) ON a.ID = al.ID
left join Receiving r WITH(NOLOCK) on a.ReceivingID = r.Id
left join Supp s WITH(NOLOCK) on a.Suppid = s.ID
left join PO_Supp_Detail psd WITH(NOLOCK) ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join PO_Supp_Detail_Spec ps WITH(NOLOCK) on psd.POID = ps.ID and psd.SEQ1 = ps.SEQ1 and psd.SEQ2 = ps.SEQ2 and ps.SpecColumnID = 'Size'
left join Pass1 q WITH(NOLOCK) on q.ID = al.WashInspector
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";

            IList<Accessory_WashExcel> listResult = ExecuteList<Accessory_WashExcel>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult.FirstOrDefault();
        }
        #endregion


        #region WashingFastness
        public Accessory_WashingFastness GetWashingFastness(Accessory_WashingFastness Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = @"
select   al.POID
		,al.ReportNo
        ,a.SCIRefno
        ,WKNo = r.ExportId
        ,a.Refno
        ,a.ArriveQty
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,Unit = psd.StockUnit
        ,Color = pc.SpecValue
        ,Size = ps.SpecValue
        ,al.Seq1
        ,al.Seq2

		,al.NonWashingFastness
		,WashingFastnessResult = al.WashingFastness
		,al.WashingFastnessEncode
		,al.WashingFastnessInspector
        ,WashingFastnessInspectorName = q.Name
		,al.WashingFastnessReceivedDate
		,al.WashingFastnessReportDate
		,al.WashingFastnessRemark
		,al.ChangeScale
		,al.ResultChange		
		,al.AcetateScale
		,al.ResultAcetate
		,al.CottonScale
		,al.ResultCotton
		,al.NylonScale
		,al.ResultNylon
		,al.PolyesterScale
		,al.ResultPolyester
		,al.AcrylicScale
		,al.ResultAcrylic		
		,al.WoolScale
		,al.ResultWool
		,al.CrossStainingScale
		,al.ResultCrossStaining
		,ali.WashingFastnessTestBeforePicture
		,ali.WashingFastnessTestAfterPicture

from AIR_Laboratory al WITH(NOLOCK)
left join SciPMSFile_AIR_Laboratory ali WITH(NOLOCK) ON ali.ID=al.ID AND  ali.POID = al.POID AND ali.Seq1 = al.Seq1 AND ali.Seq2 = al.Seq2
inner join AIR a WITH(NOLOCK) ON a.ID = al.ID
left join Receiving r WITH(NOLOCK) on a.ReceivingID = r.Id
left join Supp s WITH(NOLOCK) on a.Suppid = s.ID
left join PO_Supp_Detail psd WITH(NOLOCK) ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join PO_Supp_Detail_Spec ps WITH(NOLOCK) on psd.POID = ps.ID and psd.SEQ1 = ps.SEQ1 and psd.SEQ2 = ps.SEQ2 and ps.SpecColumnID = 'Size'
left join Pass1 q WITH(NOLOCK) on q.ID = al.WashingFastnessInspector
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";
            IList<Accessory_WashingFastness> listResult = ExecuteList<Accessory_WashingFastness>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult.FirstOrDefault();
        }

        public Accessory_WashingFastness WashingFastness_EncodeCheck(Accessory_WashingFastness Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = @"
select   1
from AIR_Laboratory
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
    AND (ResultChange = '' OR ResultAcetate = '' OR ResultCotton = '' OR ResultNylon = '' OR ResultPolyester = '' OR ResultAcrylic = '' OR ResultWool= '' OR ResultCrossStaining = '')
";
            IList<Accessory_WashingFastness> listResult = ExecuteList<Accessory_WashingFastness>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count > 0)
            {
                throw new Exception("Result cannot be empty.");
            }

            return listResult.FirstOrDefault();
        }
        public int UpdateWashingFastness(Accessory_WashingFastness Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string updateCol = string.Empty;
            string updatePicCol = string.Empty;
            string NewReportNo = GetID(Req.MDivisionID + "AT", "AIR_Laboratory", DateTime.Today, 2, "ReportNo");
            listPar.Add("@ReportNo", NewReportNo);

            #region 需要UPDATE的欄位

            if (!string.IsNullOrEmpty(Req.WashingFastnessInspector))
            {
                updateCol += $@" , WashingFastnessInspector = @WashingFastnessInspector" + Environment.NewLine;
                listPar.Add("@WashingFastnessInspector", Req.WashingFastnessInspector);
            }

            updateCol += $@" , WashingFastnessReceivedDate = @WashingFastnessReceivedDate" + Environment.NewLine;
            listPar.Add("@WashingFastnessReceivedDate", Req.WashingFastnessReceivedDate);

            updateCol += $@" , WashingFastnessReportDate = @WashingFastnessReportDate" + Environment.NewLine;
            listPar.Add("@WashingFastnessReportDate", Req.WashingFastnessReportDate);

            updateCol += $@" , WashingFastnessRemark = @WashingFastnessRemark" + Environment.NewLine;
            listPar.Add("@WashingFastnessRemark", Req.WashingFastnessRemark ?? string.Empty);

            updateCol += $@" , ChangeScale = @ChangeScale" + Environment.NewLine;
            listPar.Add("@ChangeScale", DbType.String, Req.ChangeScale ?? string.Empty);

            updateCol += $@" , ResultChange = @ResultChange" + Environment.NewLine;
            listPar.Add("@ResultChange", DbType.String, Req.ResultChange ?? string.Empty);

            updateCol += $@" , AcetateScale = @AcetateScale" + Environment.NewLine;
            listPar.Add("@AcetateScale", DbType.String, Req.AcetateScale ?? string.Empty);

            updateCol += $@" , ResultAcetate = @ResultAcetate" + Environment.NewLine;
            listPar.Add("@ResultAcetate", DbType.String, Req.ResultAcetate ?? string.Empty);

            updateCol += $@" , CottonScale = @CottonScale" + Environment.NewLine;
            listPar.Add("@CottonScale", DbType.String, Req.CottonScale ?? string.Empty);

            updateCol += $@" , ResultCotton = @ResultCotton" + Environment.NewLine;
            listPar.Add("@ResultCotton", DbType.String, Req.ResultCotton ?? string.Empty);

            updateCol += $@" , NylonScale = @NylonScale" + Environment.NewLine;
            listPar.Add("@NylonScale", DbType.String, Req.NylonScale ?? string.Empty);

            updateCol += $@" , ResultNylon = @ResultNylon" + Environment.NewLine;
            listPar.Add("@ResultNylon", Req.ResultNylon ?? string.Empty);

            updateCol += $@" , PolyesterScale = @PolyesterScale" + Environment.NewLine;
            listPar.Add("@PolyesterScale", DbType.String, Req.PolyesterScale ?? string.Empty);

            updateCol += $@" , ResultPolyester = @ResultPolyester" + Environment.NewLine;
            listPar.Add("@ResultPolyester", Req.ResultPolyester ?? string.Empty);

            updateCol += $@" , AcrylicScale = @AcrylicScale" + Environment.NewLine;
            listPar.Add("@AcrylicScale", DbType.String, Req.AcrylicScale ?? string.Empty);

            updateCol += $@" , ResultAcrylic = @ResultAcrylic" + Environment.NewLine;
            listPar.Add("@ResultAcrylic", Req.ResultAcrylic ?? string.Empty);

            updateCol += $@" , WoolScale = @WoolScale" + Environment.NewLine;
            listPar.Add("@WoolScale", DbType.String, Req.WoolScale ?? string.Empty);

            updateCol += $@" , ResultWool = @ResultWool" + Environment.NewLine;
            listPar.Add("@ResultWool", Req.ResultWool ?? string.Empty);

            updateCol += $@" , CrossStainingScale = @CrossStainingScale" + Environment.NewLine;
            listPar.Add("@CrossStainingScale", DbType.String, Req.CrossStainingScale ?? string.Empty);

            updateCol += $@" , ResultCrossStaining = @ResultCrossStaining" + Environment.NewLine;
            listPar.Add("@ResultCrossStaining", Req.ResultCrossStaining ?? string.Empty);



            updateCol += $@" , WashingFastness = (
                                                    CASE WHEN   @ResultChange = 'Pass' AND @ResultAcetate = 'Pass' AND @ResultCotton = 'Pass' AND @ResultNylon = 'Pass' 
                                                            AND @ResultPolyester = 'Pass' AND @ResultAcrylic = 'Pass' AND @ResultWool= 'Pass' AND @ResultCrossStaining = 'Pass' THEN 'Pass'
                                                        WHEN    @ResultChange = 'Fail' OR @ResultAcetate = 'Fail' OR @ResultCotton = 'Fail' OR @ResultNylon = 'Fail' 
                                                            OR @ResultPolyester = 'Fail' OR @ResultAcrylic = 'Fail' OR @ResultWool= 'Fail' OR @ResultCrossStaining = 'Fail' THEN 'Fail'
                                                        ELSE ''
                                                    END)


" + Environment.NewLine;

            updatePicCol += $@"        ,WashingFastnessTestBeforePicture = @WashingFastnessTestBeforePicture " + Environment.NewLine;
            updatePicCol += $@"        ,WashingFastnessTestAfterPicture = @WashingFastnessTestAfterPicture " + Environment.NewLine;

            if (Req.WashingFastnessTestBeforePicture != null)
            {
                listPar.Add("@WashingFastnessTestBeforePicture", Req.WashingFastnessTestBeforePicture);
            }
            else
            {
                listPar.Add("@WashingFastnessTestBeforePicture", System.Data.SqlTypes.SqlBinary.Null);
            }

            if (Req.WashingFastnessTestAfterPicture != null)
            {
                listPar.Add("@WashingFastnessTestAfterPicture", Req.WashingFastnessTestAfterPicture);
            }
            else
            {
                listPar.Add("@WashingFastnessTestAfterPicture", System.Data.SqlTypes.SqlBinary.Null);
            }
            #endregion

            string sqlCmd = $@"
SET XACT_ABORT ON

UPDATE AIR_Laboratory
SET EditDate=GETDATE()
{updateCol}
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2

UPDATE AIR_Laboratory
set ReportNo = @ReportNo
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
    and ReportNo = ''

UPDATE SciPMSFile_AIR_Laboratory
SET POID=POID
{updatePicCol}
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
";
            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }

        public int UpdateWashingFastness_WashEncodeAmend(Accessory_WashingFastness Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string updateCol = string.Empty;
            #region 需要UPDATE的欄位

            updateCol += $@" , WashingFastnessEncode = @WashingFastnessEncode" + Environment.NewLine;
            listPar.Add("@WashingFastnessEncode", Req.WashingFastnessEncode);

            if (!string.IsNullOrEmpty(Req.EditName))
            {
                updateCol += $@" , EditName = @EditName" + Environment.NewLine;
                listPar.Add("@EditName", Req.EditName);
            }
            #endregion

            string sqlCmd = $@"
UPDATE AIR_Laboratory
SET EditDate=GETDATE()
{updateCol}
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
";


            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }


        /// <summary>
        /// 取得寄信資訊
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public DataTable GetData_WashingFastnessDataTable(Accessory_WashingFastness Req)
        {
            SQLParameterCollection objParameter = new SQLParameterCollection();
            objParameter.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            objParameter.Add("@POID", Req.POID);
            objParameter.Add("@Seq1", Req.Seq1);
            objParameter.Add("@Seq2", Req.Seq2);

            string sqlGetData = @"
select  [SP#] = al.POID
		,Style = o.StyleID
		,Brand = o.BrandID
		,Season = o.SeasonID
		,Seq = al.Seq1 + ' ' + al.Seq2
        ,[WK#] = r.ExportId
		,[Arrive W/H Date]=  convert(varchar, r.WhseArrival , 111) 
        ,a.SCIRefno
        ,a.Refno
        ,Color = pc.SpecValue
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,a.ArriveQty
		,[Washing Fastness Result]=al.WashingFastness
        ,[Washing Fastness Last Test Date]=  convert(varchar, al.WashingFastnessReportDate , 111)  
		,[Washing Fastness Lab Tech	AIR_Laboratory]=al.WashingFastnessInspector
        ,Remark = al.WashingFastnessRemark
	    -- , [TestBeforePicture] = ali.WashingFastnessTestBeforePicture
	    -- , [TestAfterPicture] = ali.WashingFastnessTestAfterPicture
from AIR_Laboratory al WITH(NOLOCK)
inner join AIR a WITH(NOLOCK) ON a.ID = al.ID
INNER JOIn Orders o WITH(NOLOCK) ON o.ID = a.POID
left join SciPMSFile_AIR_Laboratory ali WITH(NOLOCK) ON ali.ID=al.ID AND ali.POID = al.POID AND ali.Seq1 = al.Seq1 AND ali.Seq2 = al.Seq2
left join Receiving r WITH(NOLOCK) on a.ReceivingID = r.Id
left join Supp s WITH(NOLOCK) on a.Suppid = s.ID
left join PO_Supp_Detail psd WITH(NOLOCK) ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);
        }

        /// <summary>
        /// 取得匯出報表資訊
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public Accessory_WashingFastnessExcel GetWashingFastnessExcel(Accessory_WashingFastness Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = @"
select   al.WashingFastnessReceivedDate
        ,psd.FactoryID
        ,al.WashingFastnessReportDate
		,o.StyleID
        ,Article = att.val
        ,a.Refno
		,o.SeasonID
        ,Color = pc.SpecValue
        ,al.ChangeScale
        ,al.ResultChange
        ,al.AcetateScale
        ,al.ResultAcetate
        ,al.CottonScale
        ,al.ResultCotton
        ,al.NylonScale
        ,al.ResultNylon
        ,al.PolyesterScale
        ,al.ResultPolyester
        ,al.AcrylicScale
        ,al.ResultAcrylic
        ,al.WoolScale
        ,al.ResultWool
        ,al.CrossStainingScale
        ,al.ResultCrossStaining
        ,Conclusions = CASE WHEN al.ResultChange = 'Fail' OR al.ResultAcetate = 'Fail' OR al.ResultCotton = 'Fail' OR al.ResultNylon = 'Fail' OR al.ResultPolyester = 'Fail' OR
                                 al.ResultAcrylic = 'Fail' OR al.ResultWool = 'Fail' OR al.ResultCrossStaining = 'Fail'  THEN 'REJECTED'
                            WHEN al.ResultChange = '' OR al.ResultAcetate = '' OR al.ResultCotton = '' OR al.ResultNylon = '' OR al.ResultPolyester = '' OR
                                 al.ResultAcrylic = '' OR al.ResultWool = '' OR al.ResultCrossStaining = ''  THEN ''
                            ELSE 'APPROVED'
                       END

		,ali.WashingFastnessTestBeforePicture
		,ali.WashingFastnessTestAfterPicture
        ,Prepared = tc.SignaturePic
        ,PreparedText = p.Name
        ,Executive = (SELECT TOP 1 SignaturePic FROM Production..Technician Where ID ='PC6000204' AND BulkAccOvenWash=1)
        ,ExecutiveText = (SELECT TOP 1 Name FROM Production..Pass1 Where ID ='PC6000204')
        ,al.ReportNo
from AIR_Laboratory al WITH(NOLOCK)
inner join Orders o WITH(NOLOCK) ON o.ID = al.POID
left join SciPMSFile_AIR_Laboratory ali WITH(NOLOCK) ON ali.ID=al.ID AND  ali.POID = al.POID AND ali.Seq1 = al.Seq1 AND ali.Seq2 = al.Seq2
inner join AIR a WITH(NOLOCK) ON a.ID = al.ID
left join PO_Supp_Detail psd WITH(NOLOCK) ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join PO_Supp_Detail_Spec pc WITH(NOLOCK) on psd.POID = pc.ID and psd.SEQ1 = pc.SEQ1 and psd.SEQ2 = pc.SEQ2 and pc.SpecColumnID = 'Color'
left join Technician tc on tc.ID = al.WashingFastnessInspector AND tc.BulkAccOvenWash=1
left join Pass1 p on p.ID = al.WashingFastnessInspector 
OUTER APPLY(
    select val = stuff((
	    select DISTINCT ',' + sa.Article
	    from Style_Article sa
	    where sa.StyleUkey = o.StyleUkey
	    FOR XML PATH('')
    ),1,1,'')
)att
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";

            IList<Accessory_WashingFastnessExcel> listResult = ExecuteList<Accessory_WashingFastnessExcel>(CommandType.Text, sqlCmd, listPar);

            if (listResult.Count == 0)
            {
                throw new Exception("No data found");
            }

            return listResult.FirstOrDefault();
        }
        #endregion


        /// <summary>
        /// 更新最終檢驗結果
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public int Update_Oven_AllResult(Accessory_Oven Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = $@"
UPDATE AIR_Laboratory
    SET  Result =   CASE    WHEN NonOven = 1 AND NonWash = 1 AND NonWashingFastness = 1 THEN 'Pass' --3個都不檢驗：直接整體PPass
	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWash = 1 AND NonWashingFastness = 1 THEN Oven                        --只有Oven檢驗
	                    WHEN NonWash = 0 AND WashEncode = 1 AND NonOven = 1 AND NonWashingFastness = 1 THEN Wash                        --只有Wash檢驗
	                    WHEN NonWashingFastness = 0 AND WashingFastnessEncode = 1 AND NonOven = 1 AND NonWash = 1 THEN WashingFastness  --只有WashingFastness檢驗

	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWash = 0 AND WashEncode = 1 THEN IIF( Oven = 'Pass' and Wash = 'Pass' , 'Pass' , 'Fail' )                                    --Oven + Wash要檢驗
	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWashingFastness= 0 AND WashingFastnessEncode = 1 THEN IIF( Oven = 'Pass' and WashingFastness = 'Pass' , 'Pass' , 'Fail' )    --Oven + WashingFastness要檢驗
	                    WHEN NonWash = 0 AND WashEncode = 1 AND NonWashingFastness= 0 AND WashingFastnessEncode = 1 THEN IIF( Wash = 'Pass' and WashingFastness = 'Pass' , 'Pass' , 'Fail' )    --Wash + WashingFastness要檢驗

	                    ELSE (	--3個都要檢驗，則套用以下判斷
			                    CASE WHEN OvenEncode != 1 OR WashEncode != 1 OR WashingFastnessEncode != 1 OR Oven='' OR Wash= '' OR WashingFastness = '' THEN ''   --其中一個未Encode或檢驗：  檢驗未完成，最終結果為空白
					                 WHEN Oven='Fail' OR Wash= 'Fail' OR WashingFastness = 'Fail'  THEN 'Fail'  --其中一個Fail：最終結果 Fail
			                    ELSE 'Pass'
			                    END
	                    )
                END
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
";
            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }
        /// <summary>
        /// 更新最終檢驗結果
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public int Update_Wash_AllResult(Accessory_Wash Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = $@"
UPDATE AIR_Laboratory
    SET  Result =   CASE    WHEN NonOven = 1 AND NonWash = 1 AND NonWashingFastness = 1 THEN 'Pass' --3個都不檢驗：直接整體PPass
	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWash = 1 AND NonWashingFastness = 1 THEN Oven                        --只有Oven檢驗
	                    WHEN NonWash = 0 AND WashEncode = 1 AND NonOven = 1 AND NonWashingFastness = 1 THEN Wash                        --只有Wash檢驗
	                    WHEN NonWashingFastness = 0 AND WashingFastnessEncode = 1 AND NonOven = 1 AND NonWash = 1 THEN WashingFastness  --只有WashingFastness檢驗

	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWash = 0 AND WashEncode = 1 THEN IIF( Oven = 'Pass' and Wash = 'Pass' , 'Pass' , 'Fail' )                                    --Oven + Wash要檢驗
	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWashingFastness= 0 AND WashingFastnessEncode = 1 THEN IIF( Oven = 'Pass' and WashingFastness = 'Pass' , 'Pass' , 'Fail' )    --Oven + WashingFastness要檢驗
	                    WHEN NonWash = 0 AND WashEncode = 1 AND NonWashingFastness= 0 AND WashingFastnessEncode = 1 THEN IIF( Wash = 'Pass' and WashingFastness = 'Pass' , 'Pass' , 'Fail' )    --Wash + WashingFastness要檢驗

	                    ELSE (	--3個都要檢驗，則套用以下判斷
			                    CASE WHEN OvenEncode != 1 OR WashEncode != 1 OR WashingFastnessEncode != 1 OR Oven='' OR Wash= '' OR WashingFastness = '' THEN ''   --其中一個未Encode或檢驗：  檢驗未完成，最終結果為空白
					                 WHEN Oven='Fail' OR Wash= 'Fail' OR WashingFastness = 'Fail'  THEN 'Fail'  --其中一個Fail：最終結果 Fail
			                    ELSE 'Pass'
			                    END
	                    )
                END
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
";
            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }
        /// <summary>
        /// 更新最終檢驗結果
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public int Update_WashingFastness_AllResult(Accessory_WashingFastness Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string sqlCmd = $@"
UPDATE AIR_Laboratory
    SET  Result =   CASE    WHEN NonOven = 1 AND NonWash = 1 AND NonWashingFastness = 1 THEN 'Pass' --3個都不檢驗：直接整體PPass
	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWash = 1 AND NonWashingFastness = 1 THEN Oven                        --只有Oven檢驗
	                    WHEN NonWash = 0 AND WashEncode = 1 AND NonOven = 1 AND NonWashingFastness = 1 THEN Wash                        --只有Wash檢驗
	                    WHEN NonWashingFastness = 0 AND WashingFastnessEncode = 1 AND NonOven = 1 AND NonWash = 1 THEN WashingFastness  --只有WashingFastness檢驗

	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWash = 0 AND WashEncode = 1 THEN IIF( Oven = 'Pass' and Wash = 'Pass' ,'Pass' ,'Fail' )                                    --Oven + Wash要檢驗
	                    WHEN NonOven = 0 AND OvenEncode = 1 AND NonWashingFastness= 0 AND WashingFastnessEncode = 1 THEN IIF( Oven = 'Pass' and WashingFastness = 'Pass' ,'Pass' ,'Fail' )    --Oven + WashingFastness要檢驗
	                    WHEN NonWash = 0 AND WashEncode = 1 AND NonWashingFastness= 0 AND WashingFastnessEncode = 1 THEN IIF( Wash = 'Pass' and WashingFastness = 'Pass' ,'Pass' ,'Fail' )    --Wash + WashingFastness要檢驗

	                    ELSE (	--3個都要檢驗，則套用以下判斷
			                    CASE WHEN OvenEncode != 1 OR WashEncode != 1 OR WashingFastnessEncode != 1 OR Oven='' OR Wash= '' OR WashingFastness = '' THEN ''   --其中一個未Encode或檢驗：  檢驗未完成，最終結果為空白
					                 WHEN Oven='Fail' OR Wash= 'Fail' OR WashingFastness = 'Fail'  THEN 'Fail'  --其中一個Fail：最終結果 Fail
			                    ELSE 'Pass'
			                    END
	                    )
                END
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
";
            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }

        public List<SelectListItem> GetScaleData()
        {

            string Cmd = @"
Select Text = ID ,Value = ID
from Scale WITH (NOLOCK) 
where junk!=1
";
            return ExecuteList<SelectListItem>(CommandType.Text, Cmd, new SQLParameterCollection()).ToList();
        }
    }
}
