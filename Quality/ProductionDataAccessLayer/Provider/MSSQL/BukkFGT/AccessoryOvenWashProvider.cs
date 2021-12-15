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
							 WHEN o.CutInline < (SELECT DATEADD(DAY, (SELECT MtlLeadTime from System) ,'2021-10-30')) THEN  o.CutInline
							 ELSE  (SELECT DATEADD(DAY, (SELECT MtlLeadTime from System) ,'2021-10-30'))
						END	
	,CompletionDate = IIF( p.AIRLabInspPercent = 100,CompletionDate.Val,NULL)
	,ArticlePercent = p.AIRLabInspPercent
	,MtlCmplt = p.Complete
	,Remark = p.AIRLaboratoryRemark
	,CreateBy = Concat (p.AddName, '-', c.Name, ' ', convert(varchar,  p.AddDate, 120) )
	,EditBy = Concat (p.EditName, '-', e.Name, ' ', convert(varchar,  p.EditDate, 120) )
from PO p
inner join Orders o ON o.ID = p.ID
left join Pass1 c on p.AddName = c.ID
left join Pass1 e on p.AddName = e.ID
OUTER APPLY(	
	SELECT Val = MAX(MaxDate) FROM (
		select MaxDate = MAX(OvenDate)
		from AIR_Laboratory
		where POID= p.ID
		UNION
		select MaxDate =MAX(WashDate)
		from AIR_Laboratory
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
	,r.ExportID
	,r.WhseArrival
	,a.SCIRefno
	,a.Refno
	,Supplier = Concat (a.SuppID, s.AbbEn)
	,psd.ColorID
	,psd.SizeSpec
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
	,WashResult = al.Wash
	,al.WashScale
	,al.WashDate
	,al.WashInspector
	,al.WashRemark
	,a.ReceivingID
	-----以下為藏在背後的Key值不會秀在畫面上-----
	,AIR_LaboratoryID = al.ID
	,a.Seq1
	,a.Seq2
from AIR_Laboratory al
inner join Air a ON a.ID = al.ID
left join Receiving r on a.ReceivingID = r.ID
left join Supp s ON s.ID = a.Suppid
left join  Po_Supp_Detail psd on a.POID = psd.ID and a.SEQ1=psd.SEQ1  and a.SEQ2=psd.SEQ2
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
            listPar.Add("@AIRLaboratoryRemark", Req.Remark);
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
SET  NonOven = @NonOven_{idx} ,NonWash = @NonWash_{idx}
,EditDate=GETDATE() ,EditName=@EditName
where POID = @ID
AND Seq1 = @Seq1_{idx}
AND Seq2 = @Seq2_{idx}
";

                listPar.Add($"@Seq1_{idx}", DbType.String, data.Seq1);
                listPar.Add($"@Seq2_{idx}", DbType.String, data.Seq2);
                listPar.Add($"@NonOven_{idx}",DbType.Boolean, data.NonOven);
                listPar.Add($"@NonWash_{idx}", DbType.Boolean, data.NonWash);
                idx++;
            }


            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }

        public int Update_AIR_Laboratory_AllResult(Accessory_ViewModel Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", Req.OrderID);
            listPar.Add("@AIRLaboratoryRemark", Req.Remark);
            listPar.Add("@EditName", Req.EditBy);

            string sqlCmd = string.Empty;

            int idx = 0;
            foreach (var data in Req.DataList)
            {
                sqlCmd += $@"
UPDATE AIR_Laboratory 
SET  Result =   CASE    WHEN NonOven = 1 AND NonWash = 1 THEN 'Pass' --兩個都不檢驗：直接整體PPass
	                    WHEN NonOven = 0 AND NonWash = 1 AND OvenEncode = 1 THEN Oven  --Wash不檢驗 + OvenEncode = true：  OvenResult Pass則視作整體Pass
	                    WHEN NonOven = 1 AND NonWash = 0 AND WashEncode = 1  THEN Wash--Oven不檢驗 + WashEncode = true：  WashResult Pass則視作整體Pass
	                    ELSE (	--兩個都要檢驗，則套用以下判斷
			                    CASE WHEN OvenEncode != 1 OR WashEncode != 1 OR Oven='' OR Wash= '' THEN ''   --其中一個未Encode或檢驗：  檢驗未完成，最終結果為空白
					                    WHEN Oven='Fail' OR Wash= 'Fail' THEN 'Fail'  --其中一個Fail：最終結果 Fail
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
        ,a.SCIRefno
        ,WKNo = r.ExportId
        ,a.Refno
        ,a.ArriveQty
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,Unit = psd.StockUnit
        ,Color = psd.ColorID
        ,psd.SizeSpec
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
from AIR_Laboratory al
inner join AIR a ON a.ID = al.ID
left join Receiving r on a.ReceivingID = r.Id
left join Supp s on a.Suppid = s.ID
left join PO_Supp_Detail psd ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join Pass1 q on q.ID = al.OvenInspector
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

        public int UpdateOvenTest(Accessory_Oven Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string updateCol = string.Empty;
            string updatePicCol = string.Empty;
            #region 需要UPDATE的欄位
            if (!string.IsNullOrEmpty(Req.Scale))
            {
                updateCol += $@" , OvenScale = @Scale" + Environment.NewLine;
                listPar.Add("@Scale", Req.Scale);
            }
            if (!string.IsNullOrEmpty(Req.OvenResult))
            {
                updateCol += $@" , Oven = @OvenResult" + Environment.NewLine;
                listPar.Add("@OvenResult", Req.OvenResult);
            }
            else
            {
                // OvenResult 不能改回空白
                // updateCol += $@" , Oven = '' " + Environment.NewLine;
            }
            if (!string.IsNullOrEmpty(Req.Remark))
            {
                updateCol += $@" , OvenRemark = @Remark" + Environment.NewLine;
                listPar.Add("@Remark", Req.Remark);
            }
            if (!string.IsNullOrEmpty(Req.OvenInspector))
            {
                updateCol += $@" , OvenInspector = @OvenInspector" + Environment.NewLine;
                listPar.Add("@OvenInspector", Req.OvenInspector);
            }

            if (Req.OvenDate.HasValue)
            {
                updateCol += $@" , OvenDate = @OvenDate" + Environment.NewLine;
                listPar.Add("@OvenDate", Req.OvenDate.Value);
            }

            updateCol += $@" , OvenEncode = @OvenEncode" + Environment.NewLine;
            listPar.Add("@OvenEncode", Req.OvenEncode);

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

UPDATE [ExtendServer].PMSFile.dbo.AIR_Laboratory
SET POID=POID
{updatePicCol}
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
";

            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }

        public int UpdateOvenTest_OvenEncode(Accessory_Oven Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string updateCol = string.Empty;
            #region 需要UPDATE的欄位
            if (!string.IsNullOrEmpty(Req.Scale))
            {
                updateCol += $@" , OvenScale = @Scale" + Environment.NewLine;
                listPar.Add("@Scale", Req.Scale);
            }
            if (!string.IsNullOrEmpty(Req.OvenResult))
            {
                updateCol += $@" , Oven = @OvenResult" + Environment.NewLine;
                listPar.Add("@OvenResult", Req.OvenResult);
            }
            else
            {
                updateCol += $@" , Oven = '' " + Environment.NewLine;
            }
            if (!string.IsNullOrEmpty(Req.Remark))
            {
                updateCol += $@" , OvenRemark = @Remark" + Environment.NewLine;
                listPar.Add("@Remark", Req.Remark);
            }
            if (!string.IsNullOrEmpty(Req.OvenInspector))
            {
                updateCol += $@" , OvenInspector = @OvenInspector" + Environment.NewLine;
                listPar.Add("@OvenInspector", Req.OvenInspector);
            }

            if (Req.OvenDate.HasValue)
            {
                updateCol += $@" , OvenDate = @OvenDate" + Environment.NewLine;
                listPar.Add("@OvenDate", Req.OvenDate.Value);
            }
            else
            {
                updateCol += $@" , OvenDate = NULL" + Environment.NewLine;
            }

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
        ,Color = psd.ColorID
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,a.ArriveQty
		,[Oven Result]=al.Oven
        ,[Oven Scale] = al.OvenScale
        ,[Oven Last Test Date]= convert(varchar, al.OvenDate , 111)  
		,[Oven Lab Tech	AIR_Laboratory]=al.OvenInspector
        ,Remark = al.OvenRemark
from AIR_Laboratory al
inner join AIR a ON a.ID = al.ID
INNER JOIn Orders o ON o.ID = a.POID
left join Receiving r on a.ReceivingID = r.Id
left join Supp s on a.Suppid = s.ID
left join PO_Supp_Detail psd ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
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
        ,a.SCIRefno
        ,WKNo = r.ExportId
        ,a.Refno
        ,a.ArriveQty
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,Unit = psd.StockUnit
        ,Color = psd.ColorID
        ,psd.SizeSpec
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
from AIR_Laboratory al
inner join AIR a ON a.ID = al.ID
left join Receiving r on a.ReceivingID = r.Id
left join Supp s on a.Suppid = s.ID
left join PO_Supp_Detail psd ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
left join Pass1 q on q.ID = al.WashInspector
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

            string updateCol = string.Empty;
            string updatePicCol = string.Empty;
            #region 需要UPDATE的欄位
            if (!string.IsNullOrEmpty(Req.Scale))
            {
                updateCol += $@" , WashScale = @Scale" + Environment.NewLine;
                listPar.Add("@Scale", Req.Scale);
            }
            if (!string.IsNullOrEmpty(Req.WashResult))
            {
                updateCol += $@" , Wash = @WashResult" + Environment.NewLine;
                listPar.Add("@WashResult", Req.WashResult);
            }
            else
            {
                // WashResult 不能改回空白
                // updateCol += $@" , Wash = '' " + Environment.NewLine;
            }
            if (!string.IsNullOrEmpty(Req.Remark))
            {
                updateCol += $@" , WashRemark = @Remark" + Environment.NewLine;
                listPar.Add("@Remark", Req.Remark);
            }
            if (!string.IsNullOrEmpty(Req.WashInspector))
            {
                updateCol += $@" , WashInspector = @WashInspector" + Environment.NewLine;
                listPar.Add("@WashInspector", Req.WashInspector);
            }

            if (Req.WashDate.HasValue)
            {
                updateCol += $@" , WashDate = @WashDate" + Environment.NewLine;
                listPar.Add("@WashDate", Req.WashDate.Value);
            }

            updateCol += $@" , WashEncode = @WashEncode" + Environment.NewLine;
            listPar.Add("@WashEncode", Req.WashEncode);

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


UPDATE [ExtendServer].PMSFile.dbo.AIR_Laboratory
SET POID=POID
{updatePicCol}
where   ID = @AIR_LaboratoryID
    and POID = @POID
    and Seq1 = @Seq1
    and Seq2 = @Seq2
";
            return ExecuteNonQuery(CommandType.Text, sqlCmd, listPar);
        }

        public int UpdateWashTest_WashEncode(Accessory_Wash Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@AIR_LaboratoryID", Req.AIR_LaboratoryID);
            listPar.Add("@POID", Req.POID);
            listPar.Add("@Seq1", Req.Seq1);
            listPar.Add("@Seq2", Req.Seq2);

            string updateCol = string.Empty;
            #region 需要UPDATE的欄位
            if (!string.IsNullOrEmpty(Req.Scale))
            {
                updateCol += $@" , WashScale = @Scale" + Environment.NewLine;
                listPar.Add("@Scale", Req.Scale);
            }
            if (!string.IsNullOrEmpty(Req.WashResult))
            {
                updateCol += $@" , Wash = @WashResult" + Environment.NewLine;
                listPar.Add("@WashResult", Req.WashResult);
            }
            else
            {
                updateCol += $@" , Wash = '' " + Environment.NewLine;
            }
            if (!string.IsNullOrEmpty(Req.Remark))
            {
                updateCol += $@" , WashRemark = @Remark" + Environment.NewLine;
                listPar.Add("@Remark", Req.Remark);
            }
            if (!string.IsNullOrEmpty(Req.WashInspector))
            {
                updateCol += $@" , WashInspector = @WashInspector" + Environment.NewLine;
                listPar.Add("@WashInspector", Req.WashInspector);
            }

            if (Req.WashDate.HasValue)
            {
                updateCol += $@" , WashDate = @WashDate" + Environment.NewLine;
                listPar.Add("@WashDate", Req.WashDate.Value);
            }
            else
            {
                updateCol += $@" , WashDate = NULL" + Environment.NewLine;
            }

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
        ,Color = psd.ColorID
        ,Supplier = Concat (a.SuppID, s.AbbEn)
        ,a.ArriveQty
		,[Wash Result]=al.Wash
        ,[Wash Scale] = al.WashScale
        ,[Wash Last Test Date]=  convert(varchar, al.WashDate , 111)  
		,[Wash Lab Tech	AIR_Laboratory]=al.WashInspector
        ,Remark = al.WashRemark
from AIR_Laboratory al
inner join AIR a ON a.ID = al.ID
INNER JOIn Orders o ON o.ID = a.POID
left join Receiving r on a.ReceivingID = r.Id
left join Supp s on a.Suppid = s.ID
left join PO_Supp_Detail psd ON psd.ID = al.POID AND psd.Seq1 = al.Seq1 AND psd.Seq2 = al.Seq2
where   al.ID=@AIR_LaboratoryID
    and al.POID=@POID
    and al.Seq1=@Seq1
    and al.Seq2=@Seq2
";

            return ExecuteDataTableByServiceConn(CommandType.Text, sqlGetData, objParameter);
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
    SET Result = CASE    WHEN NonOven = 1 AND NonWash = 1 THEN 'Pass' --兩個都不檢驗：直接整體PPass
	                        WHEN NonOven = 0 AND NonWash = 1 AND OvenEncode = 1 THEN Oven  --Wash不檢驗 + OvenEncode = true：  OvenResult Pass則視作整體Pass
	                        WHEN NonOven = 1 AND NonWash = 0 AND WashEncode = 1  THEN Wash--Oven不檢驗 + WashEncode = true：  WashResult Pass則視作整體Pass
	                        ELSE (	--兩個都要檢驗，則套用以下判斷
			                        CASE WHEN OvenEncode != 1 OR WashEncode != 1 OR Oven='' OR Wash= '' THEN ''   --其中一個未Encode或檢驗：  檢驗未完成，最終結果為空白
					                        WHEN Oven='Fail' OR Wash= 'Fail' THEN 'Fail'  --其中一個Fail：最終結果 Fail
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
    SET Result = CASE    WHEN NonOven = 1 AND NonWash = 1 THEN 'Pass' --兩個都不檢驗：直接整體PPass
	                    WHEN NonOven = 0 AND NonWash = 1 AND OvenEncode = 1 THEN Oven  --Wash不檢驗 + OvenEncode = true：  OvenResult Pass則視作整體Pass
	                    WHEN NonOven = 1 AND NonWash = 0 AND WashEncode = 1  THEN Wash--Oven不檢驗 + WashEncode = true：  WashResult Pass則視作整體Pass
	                    ELSE (	--兩個都要檢驗，則套用以下判斷
			                    CASE WHEN OvenEncode != 1 OR WashEncode != 1 OR Oven='' OR Wash= '' THEN ''   --其中一個未Encode或檢驗：  檢驗未完成，最終結果為空白
					                    WHEN Oven='Fail' OR Wash= 'Fail' THEN 'Fail'  --其中一個Fail：最終結果 Fail
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
