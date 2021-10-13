using System;
using System.Collections.Generic;
using System.Linq;
using System.Transactions;
using BusinessLogicLayer.Interface;
using BusinessLogicLayer.Service.FinalInspection;
using DatabaseObject;
using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel.FinalInspection;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;
using ToolKit;

namespace BusinessLogicLayer.Service
{
    public class FinalInspectionOthersService : IFinalInspectionOthersService
    {
        public IFinalInspectionProvider _FinalInspectionProvider { get; set; }
        public IFinalInspFromPMSProvider _FinalInspFromPMSProvider { get; set; }


        public Others GetOthersForInspection(string finalInspectionID)
        {
            Others others = new Others();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection =
                    _FinalInspectionProvider.GetFinalInspection(finalInspectionID);

                others.FinalInspectionID = finalInspectionID;
                others.CFA = finalInspection.CFA;
                others.ProductionStatus = finalInspection.ProductionStatus;
                //if (finalInspection.SubmitDate == null)
                //{
                //    others.InspectionResult = "On-going";
                //}
                //else
                //{
                //    others.InspectionResult = finalInspection.AcceptQty <= finalInspection.RejectQty ? "Fail" : "Pass";
                //}

                // ISP20211205l調整，事先帶出最後結果
                others.InspectionResult = finalInspection.AcceptQty < finalInspection.RejectQty && finalInspection.RejectQty > 0 ? "Fail" : "Pass";


                others.ShipmentStatus = finalInspection.AcceptQty < finalInspection.RejectQty && finalInspection.RejectQty > 0 ? "On Hold" : "Ship";

                others.OthersRemark = finalInspection.OthersRemark;

                others.ListOthersImageItem = _FinalInspectionProvider.GetOthersImageList(finalInspectionID).ToList();

            }
            catch (Exception ex)
            {
                others.Result = false;
                others.ErrorMessage = ex.ToString();
            }

            return others;
        }

        public List<OtherImage> GetOthersImage(string finalInspectionID)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            return _FinalInspectionProvider.GetOthersImageList(finalInspectionID).ToList();
        }

        public BaseResult UpdateOthersBack(Others others, string UserID)
        {
            BaseResult result = new BaseResult();
            using (TransactionScope transactionScope = new TransactionScope())
            {
                try
                {
                    _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                    DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection = new DatabaseObject.ManufacturingExecutionDB.FinalInspection();
                    finalInspection.ID = others.FinalInspectionID;
                    finalInspection.InspectionStep = "Insp-Measurement";

                    if (others.ProductionStatus != null)
                    {
                        finalInspection.ProductionStatus = others.ProductionStatus;
                    }

                    finalInspection.OthersRemark = others.OthersRemark;
                    finalInspection.CFA = UserID;

                    _FinalInspectionProvider.UpdateFinalInspectionByStep(finalInspection, "Insp-Others", UserID);
                    _FinalInspectionProvider.UpdateFinalInspection_OtherImage(others.FinalInspectionID, others.ListOthersImageItem.Select(o=>o.Image).ToList());

                    transactionScope.Complete();
                    transactionScope.Dispose();
                }
                catch (Exception ex)
                {
                    transactionScope.Dispose();
                    result.Result = false;
                    result.ErrorMessage = ex.ToString();
                }

            }
            return result;
        }

        public BaseResult UpdateOthersSubmit(Others others, string UserID)
        {
            BaseResult result = new BaseResult();

            if (others.ProductionStatus == null)
            {
                result.Result = false;
                result.ErrorMessage = "<Production Status> cannot be empty";
                return result;
            }

                using (TransactionScope transactionScope = new TransactionScope())
            {
                try
                {
                    _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                    //DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection = new DatabaseObject.ManufacturingExecutionDB.FinalInspection();


                    DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection =
                        _FinalInspectionProvider.GetFinalInspection(others.FinalInspectionID);

                    finalInspection.ID = others.FinalInspectionID;

                    if (others.ProductionStatus != null)
                    {
                        finalInspection.ProductionStatus = others.ProductionStatus;
                    }

                    finalInspection.OthersRemark = others.OthersRemark;
                    finalInspection.CFA = UserID;
                    finalInspection.ShipmentStatus = others.ShipmentStatus;
                    finalInspection.InspectionResult = finalInspection.AcceptQty < finalInspection.RejectQty && finalInspection.RejectQty > 0 ? "Fail" : "Pass";

                    _FinalInspectionProvider.UpdateFinalInspectionByStep(finalInspection, "Submit", UserID);
                    _FinalInspectionProvider.UpdateFinalInspection_OtherImage(others.FinalInspectionID, others.ListOthersImageItem.Select(o => o.Image).ToList());
                    transactionScope.Complete();
                    transactionScope.Dispose();
                }
                catch (Exception ex)
                {
                    transactionScope.Dispose();
                    result.Result = false;
                    result.ErrorMessage = ex.ToString();
                    return result;
                }
            }
            return result;
        }
    }
}
