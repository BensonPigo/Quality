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
    public class FinalInspectionAddDefectService : IFinalInspectionAddDefectService
    {
        public IFinalInspectionProvider _FinalInspectionProvider { get; set; }
        public IOrdersProvider _OrdersProvider { get; set; }
        public IFinalInspFromPMSProvider _FinalInspFromPMSProvider { get; set; }

        public AddDefect GetDefectForInspection(string finalInspectionID)
        {
            AddDefect addDefect = new AddDefect()
            {
                FinalInspectionID = finalInspectionID
            };

            try
            {
                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                DatabaseObject.ManufacturingExecutionDB.FinalInspection finalInspection =
                    _FinalInspectionProvider.GetFinalInspection(finalInspectionID);

                addDefect.FinalInspectionID = finalInspectionID;

                if (!finalInspection)
                {
                    addDefect.Result = finalInspection.Result;
                    addDefect.ErrorMessage = finalInspection.ErrorMessage;
                    return addDefect;
                }
                else
                {
                    addDefect.RejectQty = finalInspection.RejectQty;
                }

                _FinalInspFromPMSProvider = new FinalInspFromPMSProvider(Common.ManufacturingExecutionDataAccessLayer);

                addDefect.ListFinalInspectionDefectItem = _FinalInspFromPMSProvider.GetFinalInspectionDefectItems(finalInspectionID).ToList();

            }
            catch (Exception ex)
            {
                addDefect.Result = false;
                addDefect.ErrorMessage = ex.ToString();
            }

            return addDefect;
        }

        //public List<byte[]> GetDefectImage(long FinalInspection_DetailUkey)
        //{
        //    _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

        //    return _FinalInspectionProvider.GetFinalInspectionDefectImage(FinalInspection_DetailUkey).ToList();

        //}
        public List<ImageRemark> GetDetailItem(long FinalInspection_DetailUkey)
        {
            _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);

            return _FinalInspectionProvider.GetFinalInspectionDetail(FinalInspection_DetailUkey).ToList();

        }

        public BaseResult UpdateFinalInspectionDetail(AddDefect addDefect, string UserID)
        {
            BaseResult result = new BaseResult();
            
            try
            {
                var needUpdateDefects = addDefect.ListFinalInspectionDefectItem.Where(s => s.Qty > 0 || s.Ukey > 0);

                if (needUpdateDefects.Any())
                {
                    addDefect.ListFinalInspectionDefectItem = needUpdateDefects.ToList();
                }
                else
                {
                    addDefect.ListFinalInspectionDefectItem = new List<FinalInspectionDefectItem>();
                }

                _FinalInspectionProvider = new FinalInspectionProvider(Common.ManufacturingExecutionDataAccessLayer);
                _FinalInspectionProvider.UpdateFinalInspectionDetail(addDefect, UserID);
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
