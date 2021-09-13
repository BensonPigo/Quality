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
	,ArticlePercent = p.AIRInspPercent
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
	,al.Result
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

        public int Update(Accessory_ViewModel Req)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            listPar.Add("@ID", Req.OrderID);
            listPar.Add("@AIRLaboratoryRemark", Req.Remark);

            string sqlCmd = @"
UPDATE PO 
SET AIRLaboratoryRemark = @AIRLaboratoryRemark
WHERE ID = @ID
";
            int idx = 0;
            foreach (var data in Req.DataList)
            {
                sqlCmd += $@"
UPDATE AIR_Laboratory 
SET  NonOven = @NonOven_{idx} ,NonWash = @NonWash_{idx}
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
    }
}
