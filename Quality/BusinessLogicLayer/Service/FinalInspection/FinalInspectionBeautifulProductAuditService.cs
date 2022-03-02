using System;
using System.Collections.Generic;
using System.Linq;
using BusinessLogicLayer.Interface;
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
    public class FinalInspectionBeautifulProductAuditService : IFinalInspectionBeautifulProductAuditService
    {
        public IFinalInspectionProvider _FinalInspectionProvider { get; set; }
        public IOrdersProvider _OrdersProvider { get; set; }
        public IFinalInspFromPMSProvider _FinalInspFromPMSProvider { get; set; }

        public List<byte[]> GetBACriteriaImage(long FinalInspection_NonBACriteriaUkey)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            return _FinalInspectionProvider.GetBACriteriaImage(FinalInspection_NonBACriteriaUkey).ToList();
        }
        public List<ImageRemark> GetBA_DetailImage(long FinalInspection_NonBACriteriaUkey)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            return _FinalInspectionProvider.GetBA_DetailImage(FinalInspection_NonBACriteriaUkey).ToList();
        }

        public BeautifulProductAudit GetBeautifulProductAuditForInspection(string finalInspectionID)
        {
            BeautifulProductAudit beautifulProductAudit  = new BeautifulProductAudit();

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection =
                    _FinalInspectionProvider.GetFinalInspection(finalInspectionID);

                beautifulProductAudit.FinalInspectionID = finalInspectionID;

                if (!finalInspection)
                {
                    beautifulProductAudit.Result = finalInspection.Result;
                    beautifulProductAudit.ErrorMessage = finalInspection.ErrorMessage;
                    return beautifulProductAudit;
                }
                else
                {
                    beautifulProductAudit.BAQty = finalInspection.BAQty;
                    beautifulProductAudit.SampleSize = finalInspection.SampleSize;
                }

                beautifulProductAudit.ListBACriteria = _FinalInspectionProvider.GetBeautifulProductAuditForInspection(finalInspectionID).ToList();
            }
            catch (Exception ex)
            {
                beautifulProductAudit.Result = false;
                beautifulProductAudit.ErrorMessage = ex.ToString();
            }

            return beautifulProductAudit;
        }

        public BaseResult UpdateBeautifulProductAudit(BeautifulProductAudit beautifulProductAudit, string UserID)
        {
            BaseResult result = new BaseResult();
            
            try
            {
                var needUpdateItems = beautifulProductAudit.ListBACriteria.Where(s => s.Qty > 0 || s.Ukey > 0);

                if (needUpdateItems.Any())
                {
                    beautifulProductAudit.ListBACriteria = needUpdateItems.ToList();
                }
                else
                {
                    beautifulProductAudit.ListBACriteria = new List<BACriteriaItem>();
                }

                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                _FinalInspectionProvider.UpdateBeautifulProductAudit(beautifulProductAudit, UserID);
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.ToString();
            }

            return result;
        }
    }
}
