using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using ProductionDataAccessLayer.Interface;
using ADOHelper.Template.MSSQL;
using ADOHelper.Utility;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Data.SqlClient;
using DatabaseObject.ViewModel.FinalInspection;
using System.Linq;
using ToolKit;

namespace ProductionDataAccessLayer.Provider.MSSQL
{
    public class OrdersProvider : SQLDAL, IOrdersProvider
    {
        #region 底層連線
        public OrdersProvider(string ConString) : base(ConString) { }
        public OrdersProvider(SQLDataTransaction tra) : base(tra) { }
        #endregion

        #region CRUD Base
        /*回傳Order(Get) 詳細敘述如下*/
        /// <summary>
        /// 回傳Order
        /// </summary>
        /// <param name="Item">Order成員</param>
        /// <returns>回傳Order</returns>
        /// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public IList<Orders> Get(Orders Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection()
            {
                { "@ID", DbType.String, Item.ID },
            };
            SbSql.Append("SELECT" + Environment.NewLine);
            SbSql.Append("         ID" + Environment.NewLine);
            SbSql.Append("        ,BrandID" + Environment.NewLine);
            SbSql.Append("        ,ProgramID" + Environment.NewLine);
            SbSql.Append("        ,StyleID" + Environment.NewLine);
            SbSql.Append("        ,SeasonID" + Environment.NewLine);
            SbSql.Append("        ,ProjectID" + Environment.NewLine);
            SbSql.Append("        ,Category" + Environment.NewLine);
            SbSql.Append("        ,OrderTypeID" + Environment.NewLine);
            SbSql.Append("        ,BuyMonth" + Environment.NewLine);
            SbSql.Append("        ,Dest" + Environment.NewLine);
            SbSql.Append("        ,Model" + Environment.NewLine);
            SbSql.Append("        ,HsCode1" + Environment.NewLine);
            SbSql.Append("        ,HsCode2" + Environment.NewLine);
            SbSql.Append("        ,PayTermARID" + Environment.NewLine);
            SbSql.Append("        ,ShipTermID" + Environment.NewLine);
            SbSql.Append("        ,ShipModeList" + Environment.NewLine);
            SbSql.Append("        ,CdCodeID" + Environment.NewLine);
            SbSql.Append("        ,CPU" + Environment.NewLine);
            SbSql.Append("        ,Qty" + Environment.NewLine);
            SbSql.Append("        ,StyleUnit" + Environment.NewLine);
            SbSql.Append("        ,PoPrice" + Environment.NewLine);
            SbSql.Append("        ,CFMPrice" + Environment.NewLine);
            SbSql.Append("        ,CurrencyID" + Environment.NewLine);
            SbSql.Append("        ,Commission" + Environment.NewLine);
            SbSql.Append("        ,FactoryID" + Environment.NewLine);
            SbSql.Append("        ,BrandAreaCode" + Environment.NewLine);
            SbSql.Append("        ,BrandFTYCode" + Environment.NewLine);
            SbSql.Append("        ,CTNQty" + Environment.NewLine);
            SbSql.Append("        ,CustCDID" + Environment.NewLine);
            SbSql.Append("        ,CustPONo" + Environment.NewLine);
            SbSql.Append("        ,Customize1" + Environment.NewLine);
            SbSql.Append("        ,Customize2" + Environment.NewLine);
            SbSql.Append("        ,Customize3" + Environment.NewLine);
            SbSql.Append("        ,CFMDate" + Environment.NewLine);
            SbSql.Append("        ,BuyerDelivery" + Environment.NewLine);
            SbSql.Append("        ,SciDelivery" + Environment.NewLine);
            SbSql.Append("        ,SewInLine" + Environment.NewLine);
            SbSql.Append("        ,SewOffLine" + Environment.NewLine);
            SbSql.Append("        ,CutInLine" + Environment.NewLine);
            SbSql.Append("        ,CutOffLine" + Environment.NewLine);
            SbSql.Append("        ,PulloutDate" + Environment.NewLine);
            SbSql.Append("        ,CMPUnit" + Environment.NewLine);
            SbSql.Append("        ,CMPPrice" + Environment.NewLine);
            SbSql.Append("        ,CMPQDate" + Environment.NewLine);
            SbSql.Append("        ,CMPQRemark" + Environment.NewLine);
            SbSql.Append("        ,EachConsApv" + Environment.NewLine);
            SbSql.Append("        ,MnorderApv" + Environment.NewLine);
            SbSql.Append("        ,CRDDate" + Environment.NewLine);
            SbSql.Append("        ,InitialPlanDate" + Environment.NewLine);
            SbSql.Append("        ,PlanDate" + Environment.NewLine);
            SbSql.Append("        ,FirstProduction" + Environment.NewLine);
            SbSql.Append("        ,FirstProductionLock" + Environment.NewLine);
            SbSql.Append("        ,OrigBuyerDelivery" + Environment.NewLine);
            SbSql.Append("        ,ExCountry" + Environment.NewLine);
            SbSql.Append("        ,InDCDate" + Environment.NewLine);
            SbSql.Append("        ,CFMShipment" + Environment.NewLine);
            SbSql.Append("        ,PFETA" + Environment.NewLine);
            SbSql.Append("        ,PackLETA" + Environment.NewLine);
            SbSql.Append("        ,LETA" + Environment.NewLine);
            SbSql.Append("        ,MRHandle" + Environment.NewLine);
            SbSql.Append("        ,SMR" + Environment.NewLine);
            SbSql.Append("        ,ScanAndPack" + Environment.NewLine);
            SbSql.Append("        ,VasShas" + Environment.NewLine);
            SbSql.Append("        ,SpecialCust" + Environment.NewLine);
            SbSql.Append("        ,TissuePaper" + Environment.NewLine);
            SbSql.Append("        ,Junk" + Environment.NewLine);
            SbSql.Append("        ,Packing" + Environment.NewLine);
            SbSql.Append("        ,MarkFront" + Environment.NewLine);
            SbSql.Append("        ,MarkBack" + Environment.NewLine);
            SbSql.Append("        ,MarkLeft" + Environment.NewLine);
            SbSql.Append("        ,MarkRight" + Environment.NewLine);
            SbSql.Append("        ,Label" + Environment.NewLine);
            SbSql.Append("        ,OrderRemark" + Environment.NewLine);
            SbSql.Append("        ,ArtWorkCost" + Environment.NewLine);
            SbSql.Append("        ,StdCost" + Environment.NewLine);
            SbSql.Append("        ,CtnType" + Environment.NewLine);
            SbSql.Append("        ,FOCQty" + Environment.NewLine);
            SbSql.Append("        ,SMnorderApv" + Environment.NewLine);
            SbSql.Append("        ,FOC" + Environment.NewLine);
            SbSql.Append("        ,MnorderApv2" + Environment.NewLine);
            SbSql.Append("        ,Packing2" + Environment.NewLine);
            SbSql.Append("        ,SampleReason" + Environment.NewLine);
            SbSql.Append("        ,RainwearTestPassed" + Environment.NewLine);
            SbSql.Append("        ,SizeRange" + Environment.NewLine);
            SbSql.Append("        ,MTLComplete" + Environment.NewLine);
            SbSql.Append("        ,SpecialMark" + Environment.NewLine);
            SbSql.Append("        ,OutstandingRemark" + Environment.NewLine);
            SbSql.Append("        ,OutstandingInCharge" + Environment.NewLine);
            SbSql.Append("        ,OutstandingDate" + Environment.NewLine);
            SbSql.Append("        ,OutstandingReason" + Environment.NewLine);
            SbSql.Append("        ,StyleUkey" + Environment.NewLine);
            SbSql.Append("        ,POID" + Environment.NewLine);
            SbSql.Append("        ,OrderComboID" + Environment.NewLine);
            SbSql.Append("        ,IsNotRepeatOrMapping" + Environment.NewLine);
            SbSql.Append("        ,SplitOrderId" + Environment.NewLine);
            SbSql.Append("        ,FtyKPI" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append("        ,SewLine" + Environment.NewLine);
            SbSql.Append("        ,ActPulloutDate" + Environment.NewLine);
            SbSql.Append("        ,ProdSchdRemark" + Environment.NewLine);
            SbSql.Append("        ,IsForecast" + Environment.NewLine);
            SbSql.Append("        ,LocalOrder" + Environment.NewLine);
            SbSql.Append("        ,GMTClose" + Environment.NewLine);
            SbSql.Append("        ,TotalCTN" + Environment.NewLine);
            SbSql.Append("        ,ClogCTN" + Environment.NewLine);
            SbSql.Append("        ,FtyCTN" + Environment.NewLine);
            SbSql.Append("        ,PulloutComplete" + Environment.NewLine);
            SbSql.Append("        ,ReadyDate" + Environment.NewLine);
            SbSql.Append("        ,PulloutCTNQty" + Environment.NewLine);
            SbSql.Append("        ,Finished" + Environment.NewLine);
            SbSql.Append("        ,PFOrder" + Environment.NewLine);
            SbSql.Append("        ,SDPDate" + Environment.NewLine);
            SbSql.Append("        ,InspDate" + Environment.NewLine);
            SbSql.Append("        ,InspResult" + Environment.NewLine);
            SbSql.Append("        ,InspHandle" + Environment.NewLine);
            SbSql.Append("        ,KPILETA" + Environment.NewLine);
            SbSql.Append("        ,MTLETA" + Environment.NewLine);
            SbSql.Append("        ,SewETA" + Environment.NewLine);
            SbSql.Append("        ,PackETA" + Environment.NewLine);
            SbSql.Append("        ,MTLExport" + Environment.NewLine);
            SbSql.Append("        ,DoxType" + Environment.NewLine);
            SbSql.Append("        ,FtyGroup" + Environment.NewLine);
            SbSql.Append("        ,MDivisionID" + Environment.NewLine);
            SbSql.Append("        ,CutReadyDate" + Environment.NewLine);
            SbSql.Append("        ,SewRemark" + Environment.NewLine);
            SbSql.Append("        ,WhseClose" + Environment.NewLine);
            SbSql.Append("        ,SubconInSisterFty" + Environment.NewLine);
            SbSql.Append("        ,MCHandle" + Environment.NewLine);
            SbSql.Append("        ,LocalMR" + Environment.NewLine);
            SbSql.Append("        ,KPIChangeReason" + Environment.NewLine);
            SbSql.Append("        ,MDClose" + Environment.NewLine);
            SbSql.Append("        ,MDEditName" + Environment.NewLine);
            SbSql.Append("        ,MDEditDate" + Environment.NewLine);
            SbSql.Append("        ,ClogLastReceiveDate" + Environment.NewLine);
            SbSql.Append("        ,CPUFactor" + Environment.NewLine);
            SbSql.Append("        ,SizeUnit" + Environment.NewLine);
            SbSql.Append("        ,CuttingSP" + Environment.NewLine);
            SbSql.Append("        ,IsMixMarker" + Environment.NewLine);
            SbSql.Append("        ,EachConsSource" + Environment.NewLine);
            SbSql.Append("        ,KPIEachConsApprove" + Environment.NewLine);
            SbSql.Append("        ,KPICmpq" + Environment.NewLine);
            SbSql.Append("        ,KPIMNotice" + Environment.NewLine);
            SbSql.Append("        ,GMTComplete" + Environment.NewLine);
            SbSql.Append("        ,GFR" + Environment.NewLine);
            SbSql.Append("        ,CfaCTN" + Environment.NewLine);
            SbSql.Append("        ,DRYCTN" + Environment.NewLine);
            SbSql.Append("        ,PackErrCTN" + Environment.NewLine);
            SbSql.Append("        ,ForecastSampleGroup" + Environment.NewLine);
            SbSql.Append("        ,DyeingLoss" + Environment.NewLine);
            SbSql.Append("        ,SubconInType" + Environment.NewLine);
            SbSql.Append("        ,LastProductionDate" + Environment.NewLine);
            SbSql.Append("        ,EstPODD" + Environment.NewLine);
            SbSql.Append("        ,AirFreightByBrand" + Environment.NewLine);
            SbSql.Append("        ,AllowanceComboID" + Environment.NewLine);
            SbSql.Append("        ,ChangeMemoDate" + Environment.NewLine);
            SbSql.Append("        ,BuyBack" + Environment.NewLine);
            SbSql.Append("        ,BuyBackOrderID" + Environment.NewLine);
            SbSql.Append("        ,ForecastCategory" + Environment.NewLine);
            SbSql.Append("        ,OnSiteSample" + Environment.NewLine);
            SbSql.Append("        ,PulloutCmplDate" + Environment.NewLine);
            SbSql.Append("        ,NeedProduction" + Environment.NewLine);
            SbSql.Append("        ,IsBuyBack" + Environment.NewLine);
            SbSql.Append("        ,KeepPanels" + Environment.NewLine);
            SbSql.Append("        ,BuyBackReason" + Environment.NewLine);
            SbSql.Append("        ,IsBuyBackCrossArticle" + Environment.NewLine);
            SbSql.Append("        ,IsBuyBackCrossSizeCode" + Environment.NewLine);
            SbSql.Append("        ,KpiEachConsCheck" + Environment.NewLine);
            SbSql.Append("        ,NonRevenue" + Environment.NewLine);
            SbSql.Append("        ,CAB" + Environment.NewLine);
            SbSql.Append("        ,FinalDest" + Environment.NewLine);
            SbSql.Append("        ,Customer_PO" + Environment.NewLine);
            SbSql.Append("        ,AFS_STOCK_CATEGORY" + Environment.NewLine);
            SbSql.Append("        ,CMPLTDATE" + Environment.NewLine);
            SbSql.Append("        ,DelayCode" + Environment.NewLine);
            SbSql.Append("        ,DelayDesc" + Environment.NewLine);
            SbSql.Append("        ,HangerPack" + Environment.NewLine);
            SbSql.Append("        ,CDCodeNew" + Environment.NewLine);
            SbSql.Append("        ,SizeUnitWeight" + Environment.NewLine);
            SbSql.Append("FROM [Orders]" + Environment.NewLine);
            SbSql.Append("Where 1 = 1" + Environment.NewLine);
            if (!string.IsNullOrEmpty(Item.ID.ToString())) { SbSql.Append("And ID = @ID" + Environment.NewLine); }


            return ExecuteList<Orders>(CommandType.Text, SbSql.ToString(), objParameter);
        }
        /*建立Order(Create) 詳細敘述如下*/
        /// <summary>
        /// 建立Order
        /// </summary>
        /// <param name="Item">Order成員</param>
        /// <returns>回傳異動筆數</returns>
        /// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Create(Orders Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("INSERT INTO [Orders]" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         ID" + Environment.NewLine);
            SbSql.Append("        ,BrandID" + Environment.NewLine);
            SbSql.Append("        ,ProgramID" + Environment.NewLine);
            SbSql.Append("        ,StyleID" + Environment.NewLine);
            SbSql.Append("        ,SeasonID" + Environment.NewLine);
            SbSql.Append("        ,ProjectID" + Environment.NewLine);
            SbSql.Append("        ,Category" + Environment.NewLine);
            SbSql.Append("        ,OrderTypeID" + Environment.NewLine);
            SbSql.Append("        ,BuyMonth" + Environment.NewLine);
            SbSql.Append("        ,Dest" + Environment.NewLine);
            SbSql.Append("        ,Model" + Environment.NewLine);
            SbSql.Append("        ,HsCode1" + Environment.NewLine);
            SbSql.Append("        ,HsCode2" + Environment.NewLine);
            SbSql.Append("        ,PayTermARID" + Environment.NewLine);
            SbSql.Append("        ,ShipTermID" + Environment.NewLine);
            SbSql.Append("        ,ShipModeList" + Environment.NewLine);
            SbSql.Append("        ,CdCodeID" + Environment.NewLine);
            SbSql.Append("        ,CPU" + Environment.NewLine);
            SbSql.Append("        ,Qty" + Environment.NewLine);
            SbSql.Append("        ,StyleUnit" + Environment.NewLine);
            SbSql.Append("        ,PoPrice" + Environment.NewLine);
            SbSql.Append("        ,CFMPrice" + Environment.NewLine);
            SbSql.Append("        ,CurrencyID" + Environment.NewLine);
            SbSql.Append("        ,Commission" + Environment.NewLine);
            SbSql.Append("        ,FactoryID" + Environment.NewLine);
            SbSql.Append("        ,BrandAreaCode" + Environment.NewLine);
            SbSql.Append("        ,BrandFTYCode" + Environment.NewLine);
            SbSql.Append("        ,CTNQty" + Environment.NewLine);
            SbSql.Append("        ,CustCDID" + Environment.NewLine);
            SbSql.Append("        ,CustPONo" + Environment.NewLine);
            SbSql.Append("        ,Customize1" + Environment.NewLine);
            SbSql.Append("        ,Customize2" + Environment.NewLine);
            SbSql.Append("        ,Customize3" + Environment.NewLine);
            SbSql.Append("        ,CFMDate" + Environment.NewLine);
            SbSql.Append("        ,BuyerDelivery" + Environment.NewLine);
            SbSql.Append("        ,SciDelivery" + Environment.NewLine);
            SbSql.Append("        ,SewInLine" + Environment.NewLine);
            SbSql.Append("        ,SewOffLine" + Environment.NewLine);
            SbSql.Append("        ,CutInLine" + Environment.NewLine);
            SbSql.Append("        ,CutOffLine" + Environment.NewLine);
            SbSql.Append("        ,PulloutDate" + Environment.NewLine);
            SbSql.Append("        ,CMPUnit" + Environment.NewLine);
            SbSql.Append("        ,CMPPrice" + Environment.NewLine);
            SbSql.Append("        ,CMPQDate" + Environment.NewLine);
            SbSql.Append("        ,CMPQRemark" + Environment.NewLine);
            SbSql.Append("        ,EachConsApv" + Environment.NewLine);
            SbSql.Append("        ,MnorderApv" + Environment.NewLine);
            SbSql.Append("        ,CRDDate" + Environment.NewLine);
            SbSql.Append("        ,InitialPlanDate" + Environment.NewLine);
            SbSql.Append("        ,PlanDate" + Environment.NewLine);
            SbSql.Append("        ,FirstProduction" + Environment.NewLine);
            SbSql.Append("        ,FirstProductionLock" + Environment.NewLine);
            SbSql.Append("        ,OrigBuyerDelivery" + Environment.NewLine);
            SbSql.Append("        ,ExCountry" + Environment.NewLine);
            SbSql.Append("        ,InDCDate" + Environment.NewLine);
            SbSql.Append("        ,CFMShipment" + Environment.NewLine);
            SbSql.Append("        ,PFETA" + Environment.NewLine);
            SbSql.Append("        ,PackLETA" + Environment.NewLine);
            SbSql.Append("        ,LETA" + Environment.NewLine);
            SbSql.Append("        ,MRHandle" + Environment.NewLine);
            SbSql.Append("        ,SMR" + Environment.NewLine);
            SbSql.Append("        ,ScanAndPack" + Environment.NewLine);
            SbSql.Append("        ,VasShas" + Environment.NewLine);
            SbSql.Append("        ,SpecialCust" + Environment.NewLine);
            SbSql.Append("        ,TissuePaper" + Environment.NewLine);
            SbSql.Append("        ,Junk" + Environment.NewLine);
            SbSql.Append("        ,Packing" + Environment.NewLine);
            SbSql.Append("        ,MarkFront" + Environment.NewLine);
            SbSql.Append("        ,MarkBack" + Environment.NewLine);
            SbSql.Append("        ,MarkLeft" + Environment.NewLine);
            SbSql.Append("        ,MarkRight" + Environment.NewLine);
            SbSql.Append("        ,Label" + Environment.NewLine);
            SbSql.Append("        ,OrderRemark" + Environment.NewLine);
            SbSql.Append("        ,ArtWorkCost" + Environment.NewLine);
            SbSql.Append("        ,StdCost" + Environment.NewLine);
            SbSql.Append("        ,CtnType" + Environment.NewLine);
            SbSql.Append("        ,FOCQty" + Environment.NewLine);
            SbSql.Append("        ,SMnorderApv" + Environment.NewLine);
            SbSql.Append("        ,FOC" + Environment.NewLine);
            SbSql.Append("        ,MnorderApv2" + Environment.NewLine);
            SbSql.Append("        ,Packing2" + Environment.NewLine);
            SbSql.Append("        ,SampleReason" + Environment.NewLine);
            SbSql.Append("        ,RainwearTestPassed" + Environment.NewLine);
            SbSql.Append("        ,SizeRange" + Environment.NewLine);
            SbSql.Append("        ,MTLComplete" + Environment.NewLine);
            SbSql.Append("        ,SpecialMark" + Environment.NewLine);
            SbSql.Append("        ,OutstandingRemark" + Environment.NewLine);
            SbSql.Append("        ,OutstandingInCharge" + Environment.NewLine);
            SbSql.Append("        ,OutstandingDate" + Environment.NewLine);
            SbSql.Append("        ,OutstandingReason" + Environment.NewLine);
            SbSql.Append("        ,StyleUkey" + Environment.NewLine);
            SbSql.Append("        ,POID" + Environment.NewLine);
            SbSql.Append("        ,OrderComboID" + Environment.NewLine);
            SbSql.Append("        ,IsNotRepeatOrMapping" + Environment.NewLine);
            SbSql.Append("        ,SplitOrderId" + Environment.NewLine);
            SbSql.Append("        ,FtyKPI" + Environment.NewLine);
            SbSql.Append("        ,AddName" + Environment.NewLine);
            SbSql.Append("        ,AddDate" + Environment.NewLine);
            SbSql.Append("        ,EditName" + Environment.NewLine);
            SbSql.Append("        ,EditDate" + Environment.NewLine);
            SbSql.Append("        ,SewLine" + Environment.NewLine);
            SbSql.Append("        ,ActPulloutDate" + Environment.NewLine);
            SbSql.Append("        ,ProdSchdRemark" + Environment.NewLine);
            SbSql.Append("        ,IsForecast" + Environment.NewLine);
            SbSql.Append("        ,LocalOrder" + Environment.NewLine);
            SbSql.Append("        ,GMTClose" + Environment.NewLine);
            SbSql.Append("        ,TotalCTN" + Environment.NewLine);
            SbSql.Append("        ,ClogCTN" + Environment.NewLine);
            SbSql.Append("        ,FtyCTN" + Environment.NewLine);
            SbSql.Append("        ,PulloutComplete" + Environment.NewLine);
            SbSql.Append("        ,ReadyDate" + Environment.NewLine);
            SbSql.Append("        ,PulloutCTNQty" + Environment.NewLine);
            SbSql.Append("        ,Finished" + Environment.NewLine);
            SbSql.Append("        ,PFOrder" + Environment.NewLine);
            SbSql.Append("        ,SDPDate" + Environment.NewLine);
            SbSql.Append("        ,InspDate" + Environment.NewLine);
            SbSql.Append("        ,InspResult" + Environment.NewLine);
            SbSql.Append("        ,InspHandle" + Environment.NewLine);
            SbSql.Append("        ,KPILETA" + Environment.NewLine);
            SbSql.Append("        ,MTLETA" + Environment.NewLine);
            SbSql.Append("        ,SewETA" + Environment.NewLine);
            SbSql.Append("        ,PackETA" + Environment.NewLine);
            SbSql.Append("        ,MTLExport" + Environment.NewLine);
            SbSql.Append("        ,DoxType" + Environment.NewLine);
            SbSql.Append("        ,FtyGroup" + Environment.NewLine);
            SbSql.Append("        ,MDivisionID" + Environment.NewLine);
            SbSql.Append("        ,CutReadyDate" + Environment.NewLine);
            SbSql.Append("        ,SewRemark" + Environment.NewLine);
            SbSql.Append("        ,WhseClose" + Environment.NewLine);
            SbSql.Append("        ,SubconInSisterFty" + Environment.NewLine);
            SbSql.Append("        ,MCHandle" + Environment.NewLine);
            SbSql.Append("        ,LocalMR" + Environment.NewLine);
            SbSql.Append("        ,KPIChangeReason" + Environment.NewLine);
            SbSql.Append("        ,MDClose" + Environment.NewLine);
            SbSql.Append("        ,MDEditName" + Environment.NewLine);
            SbSql.Append("        ,MDEditDate" + Environment.NewLine);
            SbSql.Append("        ,ClogLastReceiveDate" + Environment.NewLine);
            SbSql.Append("        ,CPUFactor" + Environment.NewLine);
            SbSql.Append("        ,SizeUnit" + Environment.NewLine);
            SbSql.Append("        ,CuttingSP" + Environment.NewLine);
            SbSql.Append("        ,IsMixMarker" + Environment.NewLine);
            SbSql.Append("        ,EachConsSource" + Environment.NewLine);
            SbSql.Append("        ,KPIEachConsApprove" + Environment.NewLine);
            SbSql.Append("        ,KPICmpq" + Environment.NewLine);
            SbSql.Append("        ,KPIMNotice" + Environment.NewLine);
            SbSql.Append("        ,GMTComplete" + Environment.NewLine);
            SbSql.Append("        ,GFR" + Environment.NewLine);
            SbSql.Append("        ,CfaCTN" + Environment.NewLine);
            SbSql.Append("        ,DRYCTN" + Environment.NewLine);
            SbSql.Append("        ,PackErrCTN" + Environment.NewLine);
            SbSql.Append("        ,ForecastSampleGroup" + Environment.NewLine);
            SbSql.Append("        ,DyeingLoss" + Environment.NewLine);
            SbSql.Append("        ,SubconInType" + Environment.NewLine);
            SbSql.Append("        ,LastProductionDate" + Environment.NewLine);
            SbSql.Append("        ,EstPODD" + Environment.NewLine);
            SbSql.Append("        ,AirFreightByBrand" + Environment.NewLine);
            SbSql.Append("        ,AllowanceComboID" + Environment.NewLine);
            SbSql.Append("        ,ChangeMemoDate" + Environment.NewLine);
            SbSql.Append("        ,BuyBack" + Environment.NewLine);
            SbSql.Append("        ,BuyBackOrderID" + Environment.NewLine);
            SbSql.Append("        ,ForecastCategory" + Environment.NewLine);
            SbSql.Append("        ,OnSiteSample" + Environment.NewLine);
            SbSql.Append("        ,PulloutCmplDate" + Environment.NewLine);
            SbSql.Append("        ,NeedProduction" + Environment.NewLine);
            SbSql.Append("        ,IsBuyBack" + Environment.NewLine);
            SbSql.Append("        ,KeepPanels" + Environment.NewLine);
            SbSql.Append("        ,BuyBackReason" + Environment.NewLine);
            SbSql.Append("        ,IsBuyBackCrossArticle" + Environment.NewLine);
            SbSql.Append("        ,IsBuyBackCrossSizeCode" + Environment.NewLine);
            SbSql.Append("        ,KpiEachConsCheck" + Environment.NewLine);
            SbSql.Append("        ,NonRevenue" + Environment.NewLine);
            SbSql.Append("        ,CAB" + Environment.NewLine);
            SbSql.Append("        ,FinalDest" + Environment.NewLine);
            SbSql.Append("        ,Customer_PO" + Environment.NewLine);
            SbSql.Append("        ,AFS_STOCK_CATEGORY" + Environment.NewLine);
            SbSql.Append("        ,CMPLTDATE" + Environment.NewLine);
            SbSql.Append("        ,DelayCode" + Environment.NewLine);
            SbSql.Append("        ,DelayDesc" + Environment.NewLine);
            SbSql.Append("        ,HangerPack" + Environment.NewLine);
            SbSql.Append("        ,CDCodeNew" + Environment.NewLine);
            SbSql.Append("        ,SizeUnitWeight" + Environment.NewLine);
            SbSql.Append(")" + Environment.NewLine);
            SbSql.Append("VALUES" + Environment.NewLine);
            SbSql.Append("(" + Environment.NewLine);
            SbSql.Append("         @ID"); objParameter.Add("@ID", DbType.String, Item.ID);
            SbSql.Append("        ,@BrandID"); objParameter.Add("@BrandID", DbType.String, Item.BrandID);
            SbSql.Append("        ,@ProgramID"); objParameter.Add("@ProgramID", DbType.String, Item.ProgramID);
            SbSql.Append("        ,@StyleID"); objParameter.Add("@StyleID", DbType.String, Item.StyleID);
            SbSql.Append("        ,@SeasonID"); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID);
            SbSql.Append("        ,@ProjectID"); objParameter.Add("@ProjectID", DbType.String, Item.ProjectID);
            SbSql.Append("        ,@Category"); objParameter.Add("@Category", DbType.String, Item.Category);
            SbSql.Append("        ,@OrderTypeID"); objParameter.Add("@OrderTypeID", DbType.String, Item.OrderTypeID);
            SbSql.Append("        ,@BuyMonth"); objParameter.Add("@BuyMonth", DbType.String, Item.BuyMonth);
            SbSql.Append("        ,@Dest"); objParameter.Add("@Dest", DbType.String, Item.Dest);
            SbSql.Append("        ,@Model"); objParameter.Add("@Model", DbType.String, Item.Model);
            SbSql.Append("        ,@HsCode1"); objParameter.Add("@HsCode1", DbType.String, Item.HsCode1);
            SbSql.Append("        ,@HsCode2"); objParameter.Add("@HsCode2", DbType.String, Item.HsCode2);
            SbSql.Append("        ,@PayTermARID"); objParameter.Add("@PayTermARID", DbType.String, Item.PayTermARID);
            SbSql.Append("        ,@ShipTermID"); objParameter.Add("@ShipTermID", DbType.String, Item.ShipTermID);
            SbSql.Append("        ,@ShipModeList"); objParameter.Add("@ShipModeList", DbType.String, Item.ShipModeList);
            SbSql.Append("        ,@CdCodeID"); objParameter.Add("@CdCodeID", DbType.String, Item.CdCodeID);
            SbSql.Append("        ,@CPU"); objParameter.Add("@CPU", DbType.String, Item.CPU);
            SbSql.Append("        ,@Qty"); objParameter.Add("@Qty", DbType.Int32, Item.Qty);
            SbSql.Append("        ,@StyleUnit"); objParameter.Add("@StyleUnit", DbType.String, Item.StyleUnit);
            SbSql.Append("        ,@PoPrice"); objParameter.Add("@PoPrice", DbType.String, Item.PoPrice);
            SbSql.Append("        ,@CFMPrice"); objParameter.Add("@CFMPrice", DbType.String, Item.CFMPrice);
            SbSql.Append("        ,@CurrencyID"); objParameter.Add("@CurrencyID", DbType.String, Item.CurrencyID);
            SbSql.Append("        ,@Commission"); objParameter.Add("@Commission", DbType.String, Item.Commission);
            SbSql.Append("        ,@FactoryID"); objParameter.Add("@FactoryID", DbType.String, Item.FactoryID);
            SbSql.Append("        ,@BrandAreaCode"); objParameter.Add("@BrandAreaCode", DbType.String, Item.BrandAreaCode);
            SbSql.Append("        ,@BrandFTYCode"); objParameter.Add("@BrandFTYCode", DbType.String, Item.BrandFTYCode);
            SbSql.Append("        ,@CTNQty"); objParameter.Add("@CTNQty", DbType.String, Item.CTNQty);
            SbSql.Append("        ,@CustCDID"); objParameter.Add("@CustCDID", DbType.String, Item.CustCDID);
            SbSql.Append("        ,@CustPONo"); objParameter.Add("@CustPONo", DbType.String, Item.CustPONo);
            SbSql.Append("        ,@Customize1"); objParameter.Add("@Customize1", DbType.String, Item.Customize1);
            SbSql.Append("        ,@Customize2"); objParameter.Add("@Customize2", DbType.String, Item.Customize2);
            SbSql.Append("        ,@Customize3"); objParameter.Add("@Customize3", DbType.String, Item.Customize3);
            SbSql.Append("        ,@CFMDate"); objParameter.Add("@CFMDate", DbType.String, Item.CFMDate);
            SbSql.Append("        ,@BuyerDelivery"); objParameter.Add("@BuyerDelivery", DbType.String, Item.BuyerDelivery);
            SbSql.Append("        ,@SciDelivery"); objParameter.Add("@SciDelivery", DbType.String, Item.SciDelivery);
            SbSql.Append("        ,@SewInLine"); objParameter.Add("@SewInLine", DbType.String, Item.SewInLine);
            SbSql.Append("        ,@SewOffLine"); objParameter.Add("@SewOffLine", DbType.String, Item.SewOffLine);
            SbSql.Append("        ,@CutInLine"); objParameter.Add("@CutInLine", DbType.String, Item.CutInLine);
            SbSql.Append("        ,@CutOffLine"); objParameter.Add("@CutOffLine", DbType.String, Item.CutOffLine);
            SbSql.Append("        ,@PulloutDate"); objParameter.Add("@PulloutDate", DbType.String, Item.PulloutDate);
            SbSql.Append("        ,@CMPUnit"); objParameter.Add("@CMPUnit", DbType.String, Item.CMPUnit);
            SbSql.Append("        ,@CMPPrice"); objParameter.Add("@CMPPrice", DbType.String, Item.CMPPrice);
            SbSql.Append("        ,@CMPQDate"); objParameter.Add("@CMPQDate", DbType.String, Item.CMPQDate);
            SbSql.Append("        ,@CMPQRemark"); objParameter.Add("@CMPQRemark", DbType.String, Item.CMPQRemark);
            SbSql.Append("        ,@EachConsApv"); objParameter.Add("@EachConsApv", DbType.DateTime, Item.EachConsApv);
            SbSql.Append("        ,@MnorderApv"); objParameter.Add("@MnorderApv", DbType.DateTime, Item.MnorderApv);
            SbSql.Append("        ,@CRDDate"); objParameter.Add("@CRDDate", DbType.String, Item.CRDDate);
            SbSql.Append("        ,@InitialPlanDate"); objParameter.Add("@InitialPlanDate", DbType.String, Item.InitialPlanDate);
            SbSql.Append("        ,@PlanDate"); objParameter.Add("@PlanDate", DbType.String, Item.PlanDate);
            SbSql.Append("        ,@FirstProduction"); objParameter.Add("@FirstProduction", DbType.String, Item.FirstProduction);
            SbSql.Append("        ,@FirstProductionLock"); objParameter.Add("@FirstProductionLock", DbType.String, Item.FirstProductionLock);
            SbSql.Append("        ,@OrigBuyerDelivery"); objParameter.Add("@OrigBuyerDelivery", DbType.String, Item.OrigBuyerDelivery);
            SbSql.Append("        ,@ExCountry"); objParameter.Add("@ExCountry", DbType.String, Item.ExCountry);
            SbSql.Append("        ,@InDCDate"); objParameter.Add("@InDCDate", DbType.String, Item.InDCDate);
            SbSql.Append("        ,@CFMShipment"); objParameter.Add("@CFMShipment", DbType.String, Item.CFMShipment);
            SbSql.Append("        ,@PFETA"); objParameter.Add("@PFETA", DbType.String, Item.PFETA);
            SbSql.Append("        ,@PackLETA"); objParameter.Add("@PackLETA", DbType.String, Item.PackLETA);
            SbSql.Append("        ,@LETA"); objParameter.Add("@LETA", DbType.String, Item.LETA);
            SbSql.Append("        ,@MRHandle"); objParameter.Add("@MRHandle", DbType.String, Item.MRHandle);
            SbSql.Append("        ,@SMR"); objParameter.Add("@SMR", DbType.String, Item.SMR);
            SbSql.Append("        ,@ScanAndPack"); objParameter.Add("@ScanAndPack", DbType.String, Item.ScanAndPack);
            SbSql.Append("        ,@VasShas"); objParameter.Add("@VasShas", DbType.String, Item.VasShas);
            SbSql.Append("        ,@SpecialCust"); objParameter.Add("@SpecialCust", DbType.String, Item.SpecialCust);
            SbSql.Append("        ,@TissuePaper"); objParameter.Add("@TissuePaper", DbType.String, Item.TissuePaper);
            SbSql.Append("        ,@Junk"); objParameter.Add("@Junk", DbType.String, Item.Junk);
            SbSql.Append("        ,@Packing"); objParameter.Add("@Packing", DbType.String, Item.Packing);
            SbSql.Append("        ,@MarkFront"); objParameter.Add("@MarkFront", DbType.String, Item.MarkFront);
            SbSql.Append("        ,@MarkBack"); objParameter.Add("@MarkBack", DbType.String, Item.MarkBack);
            SbSql.Append("        ,@MarkLeft"); objParameter.Add("@MarkLeft", DbType.String, Item.MarkLeft);
            SbSql.Append("        ,@MarkRight"); objParameter.Add("@MarkRight", DbType.String, Item.MarkRight);
            SbSql.Append("        ,@Label"); objParameter.Add("@Label", DbType.String, Item.Label);
            SbSql.Append("        ,@OrderRemark"); objParameter.Add("@OrderRemark", DbType.String, Item.OrderRemark);
            SbSql.Append("        ,@ArtWorkCost"); objParameter.Add("@ArtWorkCost", DbType.String, Item.ArtWorkCost);
            SbSql.Append("        ,@StdCost"); objParameter.Add("@StdCost", DbType.String, Item.StdCost);
            SbSql.Append("        ,@CtnType"); objParameter.Add("@CtnType", DbType.String, Item.CtnType);
            SbSql.Append("        ,@FOCQty"); objParameter.Add("@FOCQty", DbType.Int32, Item.FOCQty);
            SbSql.Append("        ,@SMnorderApv"); objParameter.Add("@SMnorderApv", DbType.String, Item.SMnorderApv);
            SbSql.Append("        ,@FOC"); objParameter.Add("@FOC", DbType.String, Item.FOC);
            SbSql.Append("        ,@MnorderApv2"); objParameter.Add("@MnorderApv2", DbType.DateTime, Item.MnorderApv2);
            SbSql.Append("        ,@Packing2"); objParameter.Add("@Packing2", DbType.String, Item.Packing2);
            SbSql.Append("        ,@SampleReason"); objParameter.Add("@SampleReason", DbType.String, Item.SampleReason);
            SbSql.Append("        ,@RainwearTestPassed"); objParameter.Add("@RainwearTestPassed", DbType.String, Item.RainwearTestPassed);
            SbSql.Append("        ,@SizeRange"); objParameter.Add("@SizeRange", DbType.String, Item.SizeRange);
            SbSql.Append("        ,@MTLComplete"); objParameter.Add("@MTLComplete", DbType.String, Item.MTLComplete);
            SbSql.Append("        ,@SpecialMark"); objParameter.Add("@SpecialMark", DbType.String, Item.SpecialMark);
            SbSql.Append("        ,@OutstandingRemark"); objParameter.Add("@OutstandingRemark", DbType.String, Item.OutstandingRemark);
            SbSql.Append("        ,@OutstandingInCharge"); objParameter.Add("@OutstandingInCharge", DbType.String, Item.OutstandingInCharge);
            SbSql.Append("        ,@OutstandingDate"); objParameter.Add("@OutstandingDate", DbType.DateTime, Item.OutstandingDate);
            SbSql.Append("        ,@OutstandingReason"); objParameter.Add("@OutstandingReason", DbType.String, Item.OutstandingReason);
            SbSql.Append("        ,@StyleUkey"); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey);
            SbSql.Append("        ,@POID"); objParameter.Add("@POID", DbType.String, Item.POID);
            SbSql.Append("        ,@OrderComboID"); objParameter.Add("@OrderComboID", DbType.String, Item.OrderComboID);
            SbSql.Append("        ,@IsNotRepeatOrMapping"); objParameter.Add("@IsNotRepeatOrMapping", DbType.String, Item.IsNotRepeatOrMapping);
            SbSql.Append("        ,@SplitOrderId"); objParameter.Add("@SplitOrderId", DbType.String, Item.SplitOrderId);
            SbSql.Append("        ,@FtyKPI"); objParameter.Add("@FtyKPI", DbType.DateTime, Item.FtyKPI);
            SbSql.Append("        ,@AddName"); objParameter.Add("@AddName", DbType.String, Item.AddName);
            SbSql.Append("        ,@AddDate"); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate);
            SbSql.Append("        ,@EditName"); objParameter.Add("@EditName", DbType.String, Item.EditName);
            SbSql.Append("        ,@EditDate"); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate);
            SbSql.Append("        ,@SewLine"); objParameter.Add("@SewLine", DbType.String, Item.SewLine);
            SbSql.Append("        ,@ActPulloutDate"); objParameter.Add("@ActPulloutDate", DbType.String, Item.ActPulloutDate);
            SbSql.Append("        ,@ProdSchdRemark"); objParameter.Add("@ProdSchdRemark", DbType.String, Item.ProdSchdRemark);
            SbSql.Append("        ,@IsForecast"); objParameter.Add("@IsForecast", DbType.String, Item.IsForecast);
            SbSql.Append("        ,@LocalOrder"); objParameter.Add("@LocalOrder", DbType.String, Item.LocalOrder);
            SbSql.Append("        ,@GMTClose"); objParameter.Add("@GMTClose", DbType.String, Item.GMTClose);
            SbSql.Append("        ,@TotalCTN"); objParameter.Add("@TotalCTN", DbType.Int32, Item.TotalCTN);
            SbSql.Append("        ,@ClogCTN"); objParameter.Add("@ClogCTN", DbType.Int32, Item.ClogCTN);
            SbSql.Append("        ,@FtyCTN"); objParameter.Add("@FtyCTN", DbType.Int32, Item.FtyCTN);
            SbSql.Append("        ,@PulloutComplete"); objParameter.Add("@PulloutComplete", DbType.String, Item.PulloutComplete);
            SbSql.Append("        ,@ReadyDate"); objParameter.Add("@ReadyDate", DbType.String, Item.ReadyDate);
            SbSql.Append("        ,@PulloutCTNQty"); objParameter.Add("@PulloutCTNQty", DbType.Int32, Item.PulloutCTNQty);
            SbSql.Append("        ,@Finished"); objParameter.Add("@Finished", DbType.String, Item.Finished);
            SbSql.Append("        ,@PFOrder"); objParameter.Add("@PFOrder", DbType.String, Item.PFOrder);
            SbSql.Append("        ,@SDPDate"); objParameter.Add("@SDPDate", DbType.String, Item.SDPDate);
            SbSql.Append("        ,@InspDate"); objParameter.Add("@InspDate", DbType.String, Item.InspDate);
            SbSql.Append("        ,@InspResult"); objParameter.Add("@InspResult", DbType.String, Item.InspResult);
            SbSql.Append("        ,@InspHandle"); objParameter.Add("@InspHandle", DbType.String, Item.InspHandle);
            SbSql.Append("        ,@KPILETA"); objParameter.Add("@KPILETA", DbType.String, Item.KPILETA);
            SbSql.Append("        ,@MTLETA"); objParameter.Add("@MTLETA", DbType.String, Item.MTLETA);
            SbSql.Append("        ,@SewETA"); objParameter.Add("@SewETA", DbType.String, Item.SewETA);
            SbSql.Append("        ,@PackETA"); objParameter.Add("@PackETA", DbType.String, Item.PackETA);
            SbSql.Append("        ,@MTLExport"); objParameter.Add("@MTLExport", DbType.String, Item.MTLExport);
            SbSql.Append("        ,@DoxType"); objParameter.Add("@DoxType", DbType.String, Item.DoxType);
            SbSql.Append("        ,@FtyGroup"); objParameter.Add("@FtyGroup", DbType.String, Item.FtyGroup);
            SbSql.Append("        ,@MDivisionID"); objParameter.Add("@MDivisionID", DbType.String, Item.MDivisionID);
            SbSql.Append("        ,@CutReadyDate"); objParameter.Add("@CutReadyDate", DbType.String, Item.CutReadyDate);
            SbSql.Append("        ,@SewRemark"); objParameter.Add("@SewRemark", DbType.String, Item.SewRemark);
            SbSql.Append("        ,@WhseClose"); objParameter.Add("@WhseClose", DbType.String, Item.WhseClose);
            SbSql.Append("        ,@SubconInSisterFty"); objParameter.Add("@SubconInSisterFty", DbType.String, Item.SubconInSisterFty);
            SbSql.Append("        ,@MCHandle"); objParameter.Add("@MCHandle", DbType.String, Item.MCHandle);
            SbSql.Append("        ,@LocalMR"); objParameter.Add("@LocalMR", DbType.String, Item.LocalMR);
            SbSql.Append("        ,@KPIChangeReason"); objParameter.Add("@KPIChangeReason", DbType.String, Item.KPIChangeReason);
            SbSql.Append("        ,@MDClose"); objParameter.Add("@MDClose", DbType.String, Item.MDClose);
            SbSql.Append("        ,@MDEditName"); objParameter.Add("@MDEditName", DbType.String, Item.MDEditName);
            SbSql.Append("        ,@MDEditDate"); objParameter.Add("@MDEditDate", DbType.DateTime, Item.MDEditDate);
            SbSql.Append("        ,@ClogLastReceiveDate"); objParameter.Add("@ClogLastReceiveDate", DbType.String, Item.ClogLastReceiveDate);
            SbSql.Append("        ,@CPUFactor"); objParameter.Add("@CPUFactor", DbType.String, Item.CPUFactor);
            SbSql.Append("        ,@SizeUnit"); objParameter.Add("@SizeUnit", DbType.String, Item.SizeUnit);
            SbSql.Append("        ,@CuttingSP"); objParameter.Add("@CuttingSP", DbType.String, Item.CuttingSP);
            SbSql.Append("        ,@IsMixMarker"); objParameter.Add("@IsMixMarker", DbType.Int32, Item.IsMixMarker);
            SbSql.Append("        ,@EachConsSource"); objParameter.Add("@EachConsSource", DbType.String, Item.EachConsSource);
            SbSql.Append("        ,@KPIEachConsApprove"); objParameter.Add("@KPIEachConsApprove", DbType.String, Item.KPIEachConsApprove);
            SbSql.Append("        ,@KPICmpq"); objParameter.Add("@KPICmpq", DbType.String, Item.KPICmpq);
            SbSql.Append("        ,@KPIMNotice"); objParameter.Add("@KPIMNotice", DbType.String, Item.KPIMNotice);
            SbSql.Append("        ,@GMTComplete"); objParameter.Add("@GMTComplete", DbType.String, Item.GMTComplete);
            SbSql.Append("        ,@GFR"); objParameter.Add("@GFR", DbType.String, Item.GFR);
            SbSql.Append("        ,@CfaCTN"); objParameter.Add("@CfaCTN", DbType.Int32, Item.CfaCTN);
            SbSql.Append("        ,@DRYCTN"); objParameter.Add("@DRYCTN", DbType.Int32, Item.DRYCTN);
            SbSql.Append("        ,@PackErrCTN"); objParameter.Add("@PackErrCTN", DbType.Int32, Item.PackErrCTN);
            SbSql.Append("        ,@ForecastSampleGroup"); objParameter.Add("@ForecastSampleGroup", DbType.String, Item.ForecastSampleGroup);
            SbSql.Append("        ,@DyeingLoss"); objParameter.Add("@DyeingLoss", DbType.String, Item.DyeingLoss);
            SbSql.Append("        ,@SubconInType"); objParameter.Add("@SubconInType", DbType.String, Item.SubconInType);
            SbSql.Append("        ,@LastProductionDate"); objParameter.Add("@LastProductionDate", DbType.String, Item.LastProductionDate);
            SbSql.Append("        ,@EstPODD"); objParameter.Add("@EstPODD", DbType.String, Item.EstPODD);
            SbSql.Append("        ,@AirFreightByBrand"); objParameter.Add("@AirFreightByBrand", DbType.String, Item.AirFreightByBrand);
            SbSql.Append("        ,@AllowanceComboID"); objParameter.Add("@AllowanceComboID", DbType.String, Item.AllowanceComboID);
            SbSql.Append("        ,@ChangeMemoDate"); objParameter.Add("@ChangeMemoDate", DbType.String, Item.ChangeMemoDate);
            SbSql.Append("        ,@BuyBack"); objParameter.Add("@BuyBack", DbType.String, Item.BuyBack);
            SbSql.Append("        ,@BuyBackOrderID"); objParameter.Add("@BuyBackOrderID", DbType.String, Item.BuyBackOrderID);
            SbSql.Append("        ,@ForecastCategory"); objParameter.Add("@ForecastCategory", DbType.String, Item.ForecastCategory);
            SbSql.Append("        ,@OnSiteSample"); objParameter.Add("@OnSiteSample", DbType.String, Item.OnSiteSample);
            SbSql.Append("        ,@PulloutCmplDate"); objParameter.Add("@PulloutCmplDate", DbType.String, Item.PulloutCmplDate);
            SbSql.Append("        ,@NeedProduction"); objParameter.Add("@NeedProduction", DbType.String, Item.NeedProduction);
            SbSql.Append("        ,@IsBuyBack"); objParameter.Add("@IsBuyBack", DbType.String, Item.IsBuyBack);
            SbSql.Append("        ,@KeepPanels"); objParameter.Add("@KeepPanels", DbType.String, Item.KeepPanels);
            SbSql.Append("        ,@BuyBackReason"); objParameter.Add("@BuyBackReason", DbType.String, Item.BuyBackReason);
            SbSql.Append("        ,@IsBuyBackCrossArticle"); objParameter.Add("@IsBuyBackCrossArticle", DbType.String, Item.IsBuyBackCrossArticle);
            SbSql.Append("        ,@IsBuyBackCrossSizeCode"); objParameter.Add("@IsBuyBackCrossSizeCode", DbType.String, Item.IsBuyBackCrossSizeCode);
            SbSql.Append("        ,@KpiEachConsCheck"); objParameter.Add("@KpiEachConsCheck", DbType.String, Item.KpiEachConsCheck);
            SbSql.Append("        ,@NonRevenue"); objParameter.Add("@NonRevenue", DbType.String, Item.NonRevenue);
            SbSql.Append("        ,@CAB"); objParameter.Add("@CAB", DbType.String, Item.CAB);
            SbSql.Append("        ,@FinalDest"); objParameter.Add("@FinalDest", DbType.String, Item.FinalDest);
            SbSql.Append("        ,@Customer_PO"); objParameter.Add("@Customer_PO", DbType.String, Item.Customer_PO);
            SbSql.Append("        ,@AFS_STOCK_CATEGORY"); objParameter.Add("@AFS_STOCK_CATEGORY", DbType.String, Item.AFS_STOCK_CATEGORY);
            SbSql.Append("        ,@CMPLTDATE"); objParameter.Add("@CMPLTDATE", DbType.String, Item.CMPLTDATE);
            SbSql.Append("        ,@DelayCode"); objParameter.Add("@DelayCode", DbType.String, Item.DelayCode);
            SbSql.Append("        ,@DelayDesc"); objParameter.Add("@DelayDesc", DbType.String, Item.DelayDesc);
            SbSql.Append("        ,@HangerPack"); objParameter.Add("@HangerPack", DbType.String, Item.HangerPack);
            SbSql.Append("        ,@CDCodeNew"); objParameter.Add("@CDCodeNew", DbType.String, Item.CDCodeNew);
            SbSql.Append("        ,@SizeUnitWeight"); objParameter.Add("@SizeUnitWeight", DbType.String, Item.SizeUnitWeight);
            SbSql.Append(")" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        /*更新Order(Update) 詳細敘述如下*/
        /// <summary>
        /// 更新Order
        /// </summary>
        /// <param name="Item">Order成員</param>
        /// <returns>回傳異動筆數</returns>
        /// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Update(Orders Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("UPDATE [Orders]" + Environment.NewLine);
            SbSql.Append("SET" + Environment.NewLine);
            if (Item.ID != null) { SbSql.Append("ID=@ID" + Environment.NewLine); objParameter.Add("@ID", DbType.String, Item.ID); }
            if (Item.BrandID != null) { SbSql.Append(",BrandID=@BrandID" + Environment.NewLine); objParameter.Add("@BrandID", DbType.String, Item.BrandID); }
            if (Item.ProgramID != null) { SbSql.Append(",ProgramID=@ProgramID" + Environment.NewLine); objParameter.Add("@ProgramID", DbType.String, Item.ProgramID); }
            if (Item.StyleID != null) { SbSql.Append(",StyleID=@StyleID" + Environment.NewLine); objParameter.Add("@StyleID", DbType.String, Item.StyleID); }
            if (Item.SeasonID != null) { SbSql.Append(",SeasonID=@SeasonID" + Environment.NewLine); objParameter.Add("@SeasonID", DbType.String, Item.SeasonID); }
            if (Item.ProjectID != null) { SbSql.Append(",ProjectID=@ProjectID" + Environment.NewLine); objParameter.Add("@ProjectID", DbType.String, Item.ProjectID); }
            if (Item.Category != null) { SbSql.Append(",Category=@Category" + Environment.NewLine); objParameter.Add("@Category", DbType.String, Item.Category); }
            if (Item.OrderTypeID != null) { SbSql.Append(",OrderTypeID=@OrderTypeID" + Environment.NewLine); objParameter.Add("@OrderTypeID", DbType.String, Item.OrderTypeID); }
            if (Item.BuyMonth != null) { SbSql.Append(",BuyMonth=@BuyMonth" + Environment.NewLine); objParameter.Add("@BuyMonth", DbType.String, Item.BuyMonth); }
            if (Item.Dest != null) { SbSql.Append(",Dest=@Dest" + Environment.NewLine); objParameter.Add("@Dest", DbType.String, Item.Dest); }
            if (Item.Model != null) { SbSql.Append(",Model=@Model" + Environment.NewLine); objParameter.Add("@Model", DbType.String, Item.Model); }
            if (Item.HsCode1 != null) { SbSql.Append(",HsCode1=@HsCode1" + Environment.NewLine); objParameter.Add("@HsCode1", DbType.String, Item.HsCode1); }
            if (Item.HsCode2 != null) { SbSql.Append(",HsCode2=@HsCode2" + Environment.NewLine); objParameter.Add("@HsCode2", DbType.String, Item.HsCode2); }
            if (Item.PayTermARID != null) { SbSql.Append(",PayTermARID=@PayTermARID" + Environment.NewLine); objParameter.Add("@PayTermARID", DbType.String, Item.PayTermARID); }
            if (Item.ShipTermID != null) { SbSql.Append(",ShipTermID=@ShipTermID" + Environment.NewLine); objParameter.Add("@ShipTermID", DbType.String, Item.ShipTermID); }
            if (Item.ShipModeList != null) { SbSql.Append(",ShipModeList=@ShipModeList" + Environment.NewLine); objParameter.Add("@ShipModeList", DbType.String, Item.ShipModeList); }
            if (Item.CdCodeID != null) { SbSql.Append(",CdCodeID=@CdCodeID" + Environment.NewLine); objParameter.Add("@CdCodeID", DbType.String, Item.CdCodeID); }
            if (Item.CPU != null) { SbSql.Append(",CPU=@CPU" + Environment.NewLine); objParameter.Add("@CPU", DbType.String, Item.CPU); }
            if (Item.Qty != null) { SbSql.Append(",Qty=@Qty" + Environment.NewLine); objParameter.Add("@Qty", DbType.Int32, Item.Qty); }
            if (Item.StyleUnit != null) { SbSql.Append(",StyleUnit=@StyleUnit" + Environment.NewLine); objParameter.Add("@StyleUnit", DbType.String, Item.StyleUnit); }
            if (Item.PoPrice != null) { SbSql.Append(",PoPrice=@PoPrice" + Environment.NewLine); objParameter.Add("@PoPrice", DbType.String, Item.PoPrice); }
            if (Item.CFMPrice != null) { SbSql.Append(",CFMPrice=@CFMPrice" + Environment.NewLine); objParameter.Add("@CFMPrice", DbType.String, Item.CFMPrice); }
            if (Item.CurrencyID != null) { SbSql.Append(",CurrencyID=@CurrencyID" + Environment.NewLine); objParameter.Add("@CurrencyID", DbType.String, Item.CurrencyID); }
            if (Item.Commission != null) { SbSql.Append(",Commission=@Commission" + Environment.NewLine); objParameter.Add("@Commission", DbType.String, Item.Commission); }
            if (Item.FactoryID != null) { SbSql.Append(",FactoryID=@FactoryID" + Environment.NewLine); objParameter.Add("@FactoryID", DbType.String, Item.FactoryID); }
            if (Item.BrandAreaCode != null) { SbSql.Append(",BrandAreaCode=@BrandAreaCode" + Environment.NewLine); objParameter.Add("@BrandAreaCode", DbType.String, Item.BrandAreaCode); }
            if (Item.BrandFTYCode != null) { SbSql.Append(",BrandFTYCode=@BrandFTYCode" + Environment.NewLine); objParameter.Add("@BrandFTYCode", DbType.String, Item.BrandFTYCode); }
            if (Item.CTNQty != null) { SbSql.Append(",CTNQty=@CTNQty" + Environment.NewLine); objParameter.Add("@CTNQty", DbType.String, Item.CTNQty); }
            if (Item.CustCDID != null) { SbSql.Append(",CustCDID=@CustCDID" + Environment.NewLine); objParameter.Add("@CustCDID", DbType.String, Item.CustCDID); }
            if (Item.CustPONo != null) { SbSql.Append(",CustPONo=@CustPONo" + Environment.NewLine); objParameter.Add("@CustPONo", DbType.String, Item.CustPONo); }
            if (Item.Customize1 != null) { SbSql.Append(",Customize1=@Customize1" + Environment.NewLine); objParameter.Add("@Customize1", DbType.String, Item.Customize1); }
            if (Item.Customize2 != null) { SbSql.Append(",Customize2=@Customize2" + Environment.NewLine); objParameter.Add("@Customize2", DbType.String, Item.Customize2); }
            if (Item.Customize3 != null) { SbSql.Append(",Customize3=@Customize3" + Environment.NewLine); objParameter.Add("@Customize3", DbType.String, Item.Customize3); }
            if (Item.CFMDate != null) { SbSql.Append(",CFMDate=@CFMDate" + Environment.NewLine); objParameter.Add("@CFMDate", DbType.String, Item.CFMDate); }
            if (Item.BuyerDelivery != null) { SbSql.Append(",BuyerDelivery=@BuyerDelivery" + Environment.NewLine); objParameter.Add("@BuyerDelivery", DbType.String, Item.BuyerDelivery); }
            if (Item.SciDelivery != null) { SbSql.Append(",SciDelivery=@SciDelivery" + Environment.NewLine); objParameter.Add("@SciDelivery", DbType.String, Item.SciDelivery); }
            if (Item.SewInLine != null) { SbSql.Append(",SewInLine=@SewInLine" + Environment.NewLine); objParameter.Add("@SewInLine", DbType.String, Item.SewInLine); }
            if (Item.SewOffLine != null) { SbSql.Append(",SewOffLine=@SewOffLine" + Environment.NewLine); objParameter.Add("@SewOffLine", DbType.String, Item.SewOffLine); }
            if (Item.CutInLine != null) { SbSql.Append(",CutInLine=@CutInLine" + Environment.NewLine); objParameter.Add("@CutInLine", DbType.String, Item.CutInLine); }
            if (Item.CutOffLine != null) { SbSql.Append(",CutOffLine=@CutOffLine" + Environment.NewLine); objParameter.Add("@CutOffLine", DbType.String, Item.CutOffLine); }
            if (Item.PulloutDate != null) { SbSql.Append(",PulloutDate=@PulloutDate" + Environment.NewLine); objParameter.Add("@PulloutDate", DbType.String, Item.PulloutDate); }
            if (Item.CMPUnit != null) { SbSql.Append(",CMPUnit=@CMPUnit" + Environment.NewLine); objParameter.Add("@CMPUnit", DbType.String, Item.CMPUnit); }
            if (Item.CMPPrice != null) { SbSql.Append(",CMPPrice=@CMPPrice" + Environment.NewLine); objParameter.Add("@CMPPrice", DbType.String, Item.CMPPrice); }
            if (Item.CMPQDate != null) { SbSql.Append(",CMPQDate=@CMPQDate" + Environment.NewLine); objParameter.Add("@CMPQDate", DbType.String, Item.CMPQDate); }
            if (Item.CMPQRemark != null) { SbSql.Append(",CMPQRemark=@CMPQRemark" + Environment.NewLine); objParameter.Add("@CMPQRemark", DbType.String, Item.CMPQRemark); }
            if (Item.EachConsApv != null) { SbSql.Append(",EachConsApv=@EachConsApv" + Environment.NewLine); objParameter.Add("@EachConsApv", DbType.DateTime, Item.EachConsApv); }
            if (Item.MnorderApv != null) { SbSql.Append(",MnorderApv=@MnorderApv" + Environment.NewLine); objParameter.Add("@MnorderApv", DbType.DateTime, Item.MnorderApv); }
            if (Item.CRDDate != null) { SbSql.Append(",CRDDate=@CRDDate" + Environment.NewLine); objParameter.Add("@CRDDate", DbType.String, Item.CRDDate); }
            if (Item.InitialPlanDate != null) { SbSql.Append(",InitialPlanDate=@InitialPlanDate" + Environment.NewLine); objParameter.Add("@InitialPlanDate", DbType.String, Item.InitialPlanDate); }
            if (Item.PlanDate != null) { SbSql.Append(",PlanDate=@PlanDate" + Environment.NewLine); objParameter.Add("@PlanDate", DbType.String, Item.PlanDate); }
            if (Item.FirstProduction != null) { SbSql.Append(",FirstProduction=@FirstProduction" + Environment.NewLine); objParameter.Add("@FirstProduction", DbType.String, Item.FirstProduction); }
            if (Item.FirstProductionLock != null) { SbSql.Append(",FirstProductionLock=@FirstProductionLock" + Environment.NewLine); objParameter.Add("@FirstProductionLock", DbType.String, Item.FirstProductionLock); }
            if (Item.OrigBuyerDelivery != null) { SbSql.Append(",OrigBuyerDelivery=@OrigBuyerDelivery" + Environment.NewLine); objParameter.Add("@OrigBuyerDelivery", DbType.String, Item.OrigBuyerDelivery); }
            if (Item.ExCountry != null) { SbSql.Append(",ExCountry=@ExCountry" + Environment.NewLine); objParameter.Add("@ExCountry", DbType.String, Item.ExCountry); }
            if (Item.InDCDate != null) { SbSql.Append(",InDCDate=@InDCDate" + Environment.NewLine); objParameter.Add("@InDCDate", DbType.String, Item.InDCDate); }
            if (Item.CFMShipment != null) { SbSql.Append(",CFMShipment=@CFMShipment" + Environment.NewLine); objParameter.Add("@CFMShipment", DbType.String, Item.CFMShipment); }
            if (Item.PFETA != null) { SbSql.Append(",PFETA=@PFETA" + Environment.NewLine); objParameter.Add("@PFETA", DbType.String, Item.PFETA); }
            if (Item.PackLETA != null) { SbSql.Append(",PackLETA=@PackLETA" + Environment.NewLine); objParameter.Add("@PackLETA", DbType.String, Item.PackLETA); }
            if (Item.LETA != null) { SbSql.Append(",LETA=@LETA" + Environment.NewLine); objParameter.Add("@LETA", DbType.String, Item.LETA); }
            if (Item.MRHandle != null) { SbSql.Append(",MRHandle=@MRHandle" + Environment.NewLine); objParameter.Add("@MRHandle", DbType.String, Item.MRHandle); }
            if (Item.SMR != null) { SbSql.Append(",SMR=@SMR" + Environment.NewLine); objParameter.Add("@SMR", DbType.String, Item.SMR); }
            if (Item.ScanAndPack != null) { SbSql.Append(",ScanAndPack=@ScanAndPack" + Environment.NewLine); objParameter.Add("@ScanAndPack", DbType.String, Item.ScanAndPack); }
            if (Item.VasShas != null) { SbSql.Append(",VasShas=@VasShas" + Environment.NewLine); objParameter.Add("@VasShas", DbType.String, Item.VasShas); }
            if (Item.SpecialCust != null) { SbSql.Append(",SpecialCust=@SpecialCust" + Environment.NewLine); objParameter.Add("@SpecialCust", DbType.String, Item.SpecialCust); }
            if (Item.TissuePaper != null) { SbSql.Append(",TissuePaper=@TissuePaper" + Environment.NewLine); objParameter.Add("@TissuePaper", DbType.String, Item.TissuePaper); }
            if (Item.Junk != null) { SbSql.Append(",Junk=@Junk" + Environment.NewLine); objParameter.Add("@Junk", DbType.String, Item.Junk); }
            if (Item.Packing != null) { SbSql.Append(",Packing=@Packing" + Environment.NewLine); objParameter.Add("@Packing", DbType.String, Item.Packing); }
            if (Item.MarkFront != null) { SbSql.Append(",MarkFront=@MarkFront" + Environment.NewLine); objParameter.Add("@MarkFront", DbType.String, Item.MarkFront); }
            if (Item.MarkBack != null) { SbSql.Append(",MarkBack=@MarkBack" + Environment.NewLine); objParameter.Add("@MarkBack", DbType.String, Item.MarkBack); }
            if (Item.MarkLeft != null) { SbSql.Append(",MarkLeft=@MarkLeft" + Environment.NewLine); objParameter.Add("@MarkLeft", DbType.String, Item.MarkLeft); }
            if (Item.MarkRight != null) { SbSql.Append(",MarkRight=@MarkRight" + Environment.NewLine); objParameter.Add("@MarkRight", DbType.String, Item.MarkRight); }
            if (Item.Label != null) { SbSql.Append(",Label=@Label" + Environment.NewLine); objParameter.Add("@Label", DbType.String, Item.Label); }
            if (Item.OrderRemark != null) { SbSql.Append(",OrderRemark=@OrderRemark" + Environment.NewLine); objParameter.Add("@OrderRemark", DbType.String, Item.OrderRemark); }
            if (Item.ArtWorkCost != null) { SbSql.Append(",ArtWorkCost=@ArtWorkCost" + Environment.NewLine); objParameter.Add("@ArtWorkCost", DbType.String, Item.ArtWorkCost); }
            if (Item.StdCost != null) { SbSql.Append(",StdCost=@StdCost" + Environment.NewLine); objParameter.Add("@StdCost", DbType.String, Item.StdCost); }
            if (Item.CtnType != null) { SbSql.Append(",CtnType=@CtnType" + Environment.NewLine); objParameter.Add("@CtnType", DbType.String, Item.CtnType); }
            if (Item.FOCQty != null) { SbSql.Append(",FOCQty=@FOCQty" + Environment.NewLine); objParameter.Add("@FOCQty", DbType.Int32, Item.FOCQty); }
            if (Item.SMnorderApv != null) { SbSql.Append(",SMnorderApv=@SMnorderApv" + Environment.NewLine); objParameter.Add("@SMnorderApv", DbType.String, Item.SMnorderApv); }
            if (Item.FOC != null) { SbSql.Append(",FOC=@FOC" + Environment.NewLine); objParameter.Add("@FOC", DbType.String, Item.FOC); }
            if (Item.MnorderApv2 != null) { SbSql.Append(",MnorderApv2=@MnorderApv2" + Environment.NewLine); objParameter.Add("@MnorderApv2", DbType.DateTime, Item.MnorderApv2); }
            if (Item.Packing2 != null) { SbSql.Append(",Packing2=@Packing2" + Environment.NewLine); objParameter.Add("@Packing2", DbType.String, Item.Packing2); }
            if (Item.SampleReason != null) { SbSql.Append(",SampleReason=@SampleReason" + Environment.NewLine); objParameter.Add("@SampleReason", DbType.String, Item.SampleReason); }
            if (Item.RainwearTestPassed != null) { SbSql.Append(",RainwearTestPassed=@RainwearTestPassed" + Environment.NewLine); objParameter.Add("@RainwearTestPassed", DbType.String, Item.RainwearTestPassed); }
            if (Item.SizeRange != null) { SbSql.Append(",SizeRange=@SizeRange" + Environment.NewLine); objParameter.Add("@SizeRange", DbType.String, Item.SizeRange); }
            if (Item.MTLComplete != null) { SbSql.Append(",MTLComplete=@MTLComplete" + Environment.NewLine); objParameter.Add("@MTLComplete", DbType.String, Item.MTLComplete); }
            if (Item.SpecialMark != null) { SbSql.Append(",SpecialMark=@SpecialMark" + Environment.NewLine); objParameter.Add("@SpecialMark", DbType.String, Item.SpecialMark); }
            if (Item.OutstandingRemark != null) { SbSql.Append(",OutstandingRemark=@OutstandingRemark" + Environment.NewLine); objParameter.Add("@OutstandingRemark", DbType.String, Item.OutstandingRemark); }
            if (Item.OutstandingInCharge != null) { SbSql.Append(",OutstandingInCharge=@OutstandingInCharge" + Environment.NewLine); objParameter.Add("@OutstandingInCharge", DbType.String, Item.OutstandingInCharge); }
            if (Item.OutstandingDate != null) { SbSql.Append(",OutstandingDate=@OutstandingDate" + Environment.NewLine); objParameter.Add("@OutstandingDate", DbType.DateTime, Item.OutstandingDate); }
            if (Item.OutstandingReason != null) { SbSql.Append(",OutstandingReason=@OutstandingReason" + Environment.NewLine); objParameter.Add("@OutstandingReason", DbType.String, Item.OutstandingReason); }
            if (Item.StyleUkey != null) { SbSql.Append(",StyleUkey=@StyleUkey" + Environment.NewLine); objParameter.Add("@StyleUkey", DbType.String, Item.StyleUkey); }
            if (Item.POID != null) { SbSql.Append(",POID=@POID" + Environment.NewLine); objParameter.Add("@POID", DbType.String, Item.POID); }
            if (Item.OrderComboID != null) { SbSql.Append(",OrderComboID=@OrderComboID" + Environment.NewLine); objParameter.Add("@OrderComboID", DbType.String, Item.OrderComboID); }
            if (Item.IsNotRepeatOrMapping != null) { SbSql.Append(",IsNotRepeatOrMapping=@IsNotRepeatOrMapping" + Environment.NewLine); objParameter.Add("@IsNotRepeatOrMapping", DbType.String, Item.IsNotRepeatOrMapping); }
            if (Item.SplitOrderId != null) { SbSql.Append(",SplitOrderId=@SplitOrderId" + Environment.NewLine); objParameter.Add("@SplitOrderId", DbType.String, Item.SplitOrderId); }
            if (Item.FtyKPI != null) { SbSql.Append(",FtyKPI=@FtyKPI" + Environment.NewLine); objParameter.Add("@FtyKPI", DbType.DateTime, Item.FtyKPI); }
            if (Item.AddName != null) { SbSql.Append(",AddName=@AddName" + Environment.NewLine); objParameter.Add("@AddName", DbType.String, Item.AddName); }
            if (Item.AddDate != null) { SbSql.Append(",AddDate=@AddDate" + Environment.NewLine); objParameter.Add("@AddDate", DbType.DateTime, Item.AddDate); }
            if (Item.EditName != null) { SbSql.Append(",EditName=@EditName" + Environment.NewLine); objParameter.Add("@EditName", DbType.String, Item.EditName); }
            if (Item.EditDate != null) { SbSql.Append(",EditDate=@EditDate" + Environment.NewLine); objParameter.Add("@EditDate", DbType.DateTime, Item.EditDate); }
            if (Item.SewLine != null) { SbSql.Append(",SewLine=@SewLine" + Environment.NewLine); objParameter.Add("@SewLine", DbType.String, Item.SewLine); }
            if (Item.ActPulloutDate != null) { SbSql.Append(",ActPulloutDate=@ActPulloutDate" + Environment.NewLine); objParameter.Add("@ActPulloutDate", DbType.String, Item.ActPulloutDate); }
            if (Item.ProdSchdRemark != null) { SbSql.Append(",ProdSchdRemark=@ProdSchdRemark" + Environment.NewLine); objParameter.Add("@ProdSchdRemark", DbType.String, Item.ProdSchdRemark); }
            if (Item.IsForecast != null) { SbSql.Append(",IsForecast=@IsForecast" + Environment.NewLine); objParameter.Add("@IsForecast", DbType.String, Item.IsForecast); }
            if (Item.LocalOrder != null) { SbSql.Append(",LocalOrder=@LocalOrder" + Environment.NewLine); objParameter.Add("@LocalOrder", DbType.String, Item.LocalOrder); }
            if (Item.GMTClose != null) { SbSql.Append(",GMTClose=@GMTClose" + Environment.NewLine); objParameter.Add("@GMTClose", DbType.String, Item.GMTClose); }
            if (Item.TotalCTN != null) { SbSql.Append(",TotalCTN=@TotalCTN" + Environment.NewLine); objParameter.Add("@TotalCTN", DbType.Int32, Item.TotalCTN); }
            if (Item.ClogCTN != null) { SbSql.Append(",ClogCTN=@ClogCTN" + Environment.NewLine); objParameter.Add("@ClogCTN", DbType.Int32, Item.ClogCTN); }
            if (Item.FtyCTN != null) { SbSql.Append(",FtyCTN=@FtyCTN" + Environment.NewLine); objParameter.Add("@FtyCTN", DbType.Int32, Item.FtyCTN); }
            if (Item.PulloutComplete != null) { SbSql.Append(",PulloutComplete=@PulloutComplete" + Environment.NewLine); objParameter.Add("@PulloutComplete", DbType.String, Item.PulloutComplete); }
            if (Item.ReadyDate != null) { SbSql.Append(",ReadyDate=@ReadyDate" + Environment.NewLine); objParameter.Add("@ReadyDate", DbType.String, Item.ReadyDate); }
            if (Item.PulloutCTNQty != null) { SbSql.Append(",PulloutCTNQty=@PulloutCTNQty" + Environment.NewLine); objParameter.Add("@PulloutCTNQty", DbType.Int32, Item.PulloutCTNQty); }
            if (Item.Finished != null) { SbSql.Append(",Finished=@Finished" + Environment.NewLine); objParameter.Add("@Finished", DbType.String, Item.Finished); }
            if (Item.PFOrder != null) { SbSql.Append(",PFOrder=@PFOrder" + Environment.NewLine); objParameter.Add("@PFOrder", DbType.String, Item.PFOrder); }
            if (Item.SDPDate != null) { SbSql.Append(",SDPDate=@SDPDate" + Environment.NewLine); objParameter.Add("@SDPDate", DbType.String, Item.SDPDate); }
            if (Item.InspDate != null) { SbSql.Append(",InspDate=@InspDate" + Environment.NewLine); objParameter.Add("@InspDate", DbType.String, Item.InspDate); }
            if (Item.InspResult != null) { SbSql.Append(",InspResult=@InspResult" + Environment.NewLine); objParameter.Add("@InspResult", DbType.String, Item.InspResult); }
            if (Item.InspHandle != null) { SbSql.Append(",InspHandle=@InspHandle" + Environment.NewLine); objParameter.Add("@InspHandle", DbType.String, Item.InspHandle); }
            if (Item.KPILETA != null) { SbSql.Append(",KPILETA=@KPILETA" + Environment.NewLine); objParameter.Add("@KPILETA", DbType.String, Item.KPILETA); }
            if (Item.MTLETA != null) { SbSql.Append(",MTLETA=@MTLETA" + Environment.NewLine); objParameter.Add("@MTLETA", DbType.String, Item.MTLETA); }
            if (Item.SewETA != null) { SbSql.Append(",SewETA=@SewETA" + Environment.NewLine); objParameter.Add("@SewETA", DbType.String, Item.SewETA); }
            if (Item.PackETA != null) { SbSql.Append(",PackETA=@PackETA" + Environment.NewLine); objParameter.Add("@PackETA", DbType.String, Item.PackETA); }
            if (Item.MTLExport != null) { SbSql.Append(",MTLExport=@MTLExport" + Environment.NewLine); objParameter.Add("@MTLExport", DbType.String, Item.MTLExport); }
            if (Item.DoxType != null) { SbSql.Append(",DoxType=@DoxType" + Environment.NewLine); objParameter.Add("@DoxType", DbType.String, Item.DoxType); }
            if (Item.FtyGroup != null) { SbSql.Append(",FtyGroup=@FtyGroup" + Environment.NewLine); objParameter.Add("@FtyGroup", DbType.String, Item.FtyGroup); }
            if (Item.MDivisionID != null) { SbSql.Append(",MDivisionID=@MDivisionID" + Environment.NewLine); objParameter.Add("@MDivisionID", DbType.String, Item.MDivisionID); }
            if (Item.CutReadyDate != null) { SbSql.Append(",CutReadyDate=@CutReadyDate" + Environment.NewLine); objParameter.Add("@CutReadyDate", DbType.String, Item.CutReadyDate); }
            if (Item.SewRemark != null) { SbSql.Append(",SewRemark=@SewRemark" + Environment.NewLine); objParameter.Add("@SewRemark", DbType.String, Item.SewRemark); }
            if (Item.WhseClose != null) { SbSql.Append(",WhseClose=@WhseClose" + Environment.NewLine); objParameter.Add("@WhseClose", DbType.String, Item.WhseClose); }
            if (Item.SubconInSisterFty != null) { SbSql.Append(",SubconInSisterFty=@SubconInSisterFty" + Environment.NewLine); objParameter.Add("@SubconInSisterFty", DbType.String, Item.SubconInSisterFty); }
            if (Item.MCHandle != null) { SbSql.Append(",MCHandle=@MCHandle" + Environment.NewLine); objParameter.Add("@MCHandle", DbType.String, Item.MCHandle); }
            if (Item.LocalMR != null) { SbSql.Append(",LocalMR=@LocalMR" + Environment.NewLine); objParameter.Add("@LocalMR", DbType.String, Item.LocalMR); }
            if (Item.KPIChangeReason != null) { SbSql.Append(",KPIChangeReason=@KPIChangeReason" + Environment.NewLine); objParameter.Add("@KPIChangeReason", DbType.String, Item.KPIChangeReason); }
            if (Item.MDClose != null) { SbSql.Append(",MDClose=@MDClose" + Environment.NewLine); objParameter.Add("@MDClose", DbType.String, Item.MDClose); }
            if (Item.MDEditName != null) { SbSql.Append(",MDEditName=@MDEditName" + Environment.NewLine); objParameter.Add("@MDEditName", DbType.String, Item.MDEditName); }
            if (Item.MDEditDate != null) { SbSql.Append(",MDEditDate=@MDEditDate" + Environment.NewLine); objParameter.Add("@MDEditDate", DbType.DateTime, Item.MDEditDate); }
            if (Item.ClogLastReceiveDate != null) { SbSql.Append(",ClogLastReceiveDate=@ClogLastReceiveDate" + Environment.NewLine); objParameter.Add("@ClogLastReceiveDate", DbType.String, Item.ClogLastReceiveDate); }
            if (Item.CPUFactor != null) { SbSql.Append(",CPUFactor=@CPUFactor" + Environment.NewLine); objParameter.Add("@CPUFactor", DbType.String, Item.CPUFactor); }
            if (Item.SizeUnit != null) { SbSql.Append(",SizeUnit=@SizeUnit" + Environment.NewLine); objParameter.Add("@SizeUnit", DbType.String, Item.SizeUnit); }
            if (Item.CuttingSP != null) { SbSql.Append(",CuttingSP=@CuttingSP" + Environment.NewLine); objParameter.Add("@CuttingSP", DbType.String, Item.CuttingSP); }
            if (Item.IsMixMarker != null) { SbSql.Append(",IsMixMarker=@IsMixMarker" + Environment.NewLine); objParameter.Add("@IsMixMarker", DbType.Int32, Item.IsMixMarker); }
            if (Item.EachConsSource != null) { SbSql.Append(",EachConsSource=@EachConsSource" + Environment.NewLine); objParameter.Add("@EachConsSource", DbType.String, Item.EachConsSource); }
            if (Item.KPIEachConsApprove != null) { SbSql.Append(",KPIEachConsApprove=@KPIEachConsApprove" + Environment.NewLine); objParameter.Add("@KPIEachConsApprove", DbType.String, Item.KPIEachConsApprove); }
            if (Item.KPICmpq != null) { SbSql.Append(",KPICmpq=@KPICmpq" + Environment.NewLine); objParameter.Add("@KPICmpq", DbType.String, Item.KPICmpq); }
            if (Item.KPIMNotice != null) { SbSql.Append(",KPIMNotice=@KPIMNotice" + Environment.NewLine); objParameter.Add("@KPIMNotice", DbType.String, Item.KPIMNotice); }
            if (Item.GMTComplete != null) { SbSql.Append(",GMTComplete=@GMTComplete" + Environment.NewLine); objParameter.Add("@GMTComplete", DbType.String, Item.GMTComplete); }
            if (Item.GFR != null) { SbSql.Append(",GFR=@GFR" + Environment.NewLine); objParameter.Add("@GFR", DbType.String, Item.GFR); }
            if (Item.CfaCTN != null) { SbSql.Append(",CfaCTN=@CfaCTN" + Environment.NewLine); objParameter.Add("@CfaCTN", DbType.Int32, Item.CfaCTN); }
            if (Item.DRYCTN != null) { SbSql.Append(",DRYCTN=@DRYCTN" + Environment.NewLine); objParameter.Add("@DRYCTN", DbType.Int32, Item.DRYCTN); }
            if (Item.PackErrCTN != null) { SbSql.Append(",PackErrCTN=@PackErrCTN" + Environment.NewLine); objParameter.Add("@PackErrCTN", DbType.Int32, Item.PackErrCTN); }
            if (Item.ForecastSampleGroup != null) { SbSql.Append(",ForecastSampleGroup=@ForecastSampleGroup" + Environment.NewLine); objParameter.Add("@ForecastSampleGroup", DbType.String, Item.ForecastSampleGroup); }
            if (Item.DyeingLoss != null) { SbSql.Append(",DyeingLoss=@DyeingLoss" + Environment.NewLine); objParameter.Add("@DyeingLoss", DbType.String, Item.DyeingLoss); }
            if (Item.SubconInType != null) { SbSql.Append(",SubconInType=@SubconInType" + Environment.NewLine); objParameter.Add("@SubconInType", DbType.String, Item.SubconInType); }
            if (Item.LastProductionDate != null) { SbSql.Append(",LastProductionDate=@LastProductionDate" + Environment.NewLine); objParameter.Add("@LastProductionDate", DbType.String, Item.LastProductionDate); }
            if (Item.EstPODD != null) { SbSql.Append(",EstPODD=@EstPODD" + Environment.NewLine); objParameter.Add("@EstPODD", DbType.String, Item.EstPODD); }
            if (Item.AirFreightByBrand != null) { SbSql.Append(",AirFreightByBrand=@AirFreightByBrand" + Environment.NewLine); objParameter.Add("@AirFreightByBrand", DbType.String, Item.AirFreightByBrand); }
            if (Item.AllowanceComboID != null) { SbSql.Append(",AllowanceComboID=@AllowanceComboID" + Environment.NewLine); objParameter.Add("@AllowanceComboID", DbType.String, Item.AllowanceComboID); }
            if (Item.ChangeMemoDate != null) { SbSql.Append(",ChangeMemoDate=@ChangeMemoDate" + Environment.NewLine); objParameter.Add("@ChangeMemoDate", DbType.String, Item.ChangeMemoDate); }
            if (Item.BuyBack != null) { SbSql.Append(",BuyBack=@BuyBack" + Environment.NewLine); objParameter.Add("@BuyBack", DbType.String, Item.BuyBack); }
            if (Item.BuyBackOrderID != null) { SbSql.Append(",BuyBackOrderID=@BuyBackOrderID" + Environment.NewLine); objParameter.Add("@BuyBackOrderID", DbType.String, Item.BuyBackOrderID); }
            if (Item.ForecastCategory != null) { SbSql.Append(",ForecastCategory=@ForecastCategory" + Environment.NewLine); objParameter.Add("@ForecastCategory", DbType.String, Item.ForecastCategory); }
            if (Item.OnSiteSample != null) { SbSql.Append(",OnSiteSample=@OnSiteSample" + Environment.NewLine); objParameter.Add("@OnSiteSample", DbType.String, Item.OnSiteSample); }
            if (Item.PulloutCmplDate != null) { SbSql.Append(",PulloutCmplDate=@PulloutCmplDate" + Environment.NewLine); objParameter.Add("@PulloutCmplDate", DbType.String, Item.PulloutCmplDate); }
            if (Item.NeedProduction != null) { SbSql.Append(",NeedProduction=@NeedProduction" + Environment.NewLine); objParameter.Add("@NeedProduction", DbType.String, Item.NeedProduction); }
            if (Item.IsBuyBack != null) { SbSql.Append(",IsBuyBack=@IsBuyBack" + Environment.NewLine); objParameter.Add("@IsBuyBack", DbType.String, Item.IsBuyBack); }
            if (Item.KeepPanels != null) { SbSql.Append(",KeepPanels=@KeepPanels" + Environment.NewLine); objParameter.Add("@KeepPanels", DbType.String, Item.KeepPanels); }
            if (Item.BuyBackReason != null) { SbSql.Append(",BuyBackReason=@BuyBackReason" + Environment.NewLine); objParameter.Add("@BuyBackReason", DbType.String, Item.BuyBackReason); }
            if (Item.IsBuyBackCrossArticle != null) { SbSql.Append(",IsBuyBackCrossArticle=@IsBuyBackCrossArticle" + Environment.NewLine); objParameter.Add("@IsBuyBackCrossArticle", DbType.String, Item.IsBuyBackCrossArticle); }
            if (Item.IsBuyBackCrossSizeCode != null) { SbSql.Append(",IsBuyBackCrossSizeCode=@IsBuyBackCrossSizeCode" + Environment.NewLine); objParameter.Add("@IsBuyBackCrossSizeCode", DbType.String, Item.IsBuyBackCrossSizeCode); }
            if (Item.KpiEachConsCheck != null) { SbSql.Append(",KpiEachConsCheck=@KpiEachConsCheck" + Environment.NewLine); objParameter.Add("@KpiEachConsCheck", DbType.String, Item.KpiEachConsCheck); }
            if (Item.NonRevenue != null) { SbSql.Append(",NonRevenue=@NonRevenue" + Environment.NewLine); objParameter.Add("@NonRevenue", DbType.String, Item.NonRevenue); }
            if (Item.CAB != null) { SbSql.Append(",CAB=@CAB" + Environment.NewLine); objParameter.Add("@CAB", DbType.String, Item.CAB); }
            if (Item.FinalDest != null) { SbSql.Append(",FinalDest=@FinalDest" + Environment.NewLine); objParameter.Add("@FinalDest", DbType.String, Item.FinalDest); }
            if (Item.Customer_PO != null) { SbSql.Append(",Customer_PO=@Customer_PO" + Environment.NewLine); objParameter.Add("@Customer_PO", DbType.String, Item.Customer_PO); }
            if (Item.AFS_STOCK_CATEGORY != null) { SbSql.Append(",AFS_STOCK_CATEGORY=@AFS_STOCK_CATEGORY" + Environment.NewLine); objParameter.Add("@AFS_STOCK_CATEGORY", DbType.String, Item.AFS_STOCK_CATEGORY); }
            if (Item.CMPLTDATE != null) { SbSql.Append(",CMPLTDATE=@CMPLTDATE" + Environment.NewLine); objParameter.Add("@CMPLTDATE", DbType.String, Item.CMPLTDATE); }
            if (Item.DelayCode != null) { SbSql.Append(",DelayCode=@DelayCode" + Environment.NewLine); objParameter.Add("@DelayCode", DbType.String, Item.DelayCode); }
            if (Item.DelayDesc != null) { SbSql.Append(",DelayDesc=@DelayDesc" + Environment.NewLine); objParameter.Add("@DelayDesc", DbType.String, Item.DelayDesc); }
            if (Item.HangerPack != null) { SbSql.Append(",HangerPack=@HangerPack" + Environment.NewLine); objParameter.Add("@HangerPack", DbType.String, Item.HangerPack); }
            if (Item.CDCodeNew != null) { SbSql.Append(",CDCodeNew=@CDCodeNew" + Environment.NewLine); objParameter.Add("@CDCodeNew", DbType.String, Item.CDCodeNew); }
            if (Item.SizeUnitWeight != null) { SbSql.Append(",SizeUnitWeight=@SizeUnitWeight" + Environment.NewLine); objParameter.Add("@SizeUnitWeight", DbType.String, Item.SizeUnitWeight); }
            SbSql.Append("WHERE 1 = 1" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        /*刪除Order(Delete) 詳細敘述如下*/
        /// <summary>
        /// 刪除Order
        /// </summary>
        /// <param name="Item">Order成員</param>
        /// <returns>回傳異動筆數</returns>
        /// <info>Author: Admin; Date: 2021/08/05  </info>
        /// <history>
        /// xx.  YYYY/MM/DD   Ver   Author      Comments
        /// ===  ==========  ====  ==========  ==========
        /// 01.  2021/08/05  1.00    Admin        Create
        /// </history>
        public int Delete(Orders Item)
        {
            StringBuilder SbSql = new StringBuilder();
            SQLParameterCollection objParameter = new SQLParameterCollection();
            SbSql.Append("DELETE FROM [Orders]" + Environment.NewLine);




            return ExecuteNonQuery(CommandType.Text, SbSql.ToString(), objParameter);
        }
        #endregion

        public IList<Orders> GetOrderForInspection(FinalInspection_Request requestItem)
        {
            string sqlGetData = string.Empty;
            SQLParameterCollection listPar = new SQLParameterCollection();

            string where = string.Empty;

            listPar.Add("@Ftygroup", requestItem.FactoryID);
            if (!string.IsNullOrEmpty(requestItem.SP))
            {
                listPar.Add("@SP", requestItem.SP);
                where += " and o.id = @SP ";
            }

            if (!string.IsNullOrEmpty(requestItem.POID))
            {
                listPar.Add("@POID", requestItem.POID);
                where += " and o.POID = @POID ";
            }

            if (!string.IsNullOrEmpty(requestItem.StyleID))
            {
                listPar.Add("@StyleID", requestItem.StyleID);
                where += " and o.StyleID = @StyleID ";
            }

            sqlGetData = $@"
select o.ID 
     , o.POID 
     , o.Qty  
     , o.StyleID  
     , o.SeasonID 
     , o.BrandID 
  from orders o
where o.ftygroup = @Ftygroup and o.PulloutComplete = 0
        {where}
";

            return ExecuteList<Orders>(CommandType.Text, sqlGetData, listPar);
        }

        public IList<SelectedPO> GetSelectedPOForInspection(List<string> listOrderID)
        {
            SQLParameterCollection listPar = new SQLParameterCollection();
            string whereOrderID = listOrderID.Select(s => $"'{s}'").JoinToString(",");
            string sqlGetData = $@"
select  [OrderID] = o.id,
        o.POID,
        o.StyleID,
        o.SeasonID,
        o.BrandID,
        o.Qty,
        [AvailableQty] = 0,
        [Cartons] = ''
  from orders o
 where  o.id in ({whereOrderID})
";
            return ExecuteList<SelectedPO>(CommandType.Text, sqlGetData, listPar);
        }
    }
}
