using ADOHelper.Utility;
using BusinessLogicLayer.Interface.BulkFGT;
using DatabaseObject;
using DatabaseObject.ViewModel.BulkFGT;
using MICS.DataAccessLayer.Interface;
using MICS.DataAccessLayer.Provider.MSSQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static MICS.DataAccessLayer.Provider.MSSQL.ColorFastnessDetailProvider;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class FabricColorFastness_Service : IFabricColorFastness_Service
    {
        private IColorFastnessProvider _IColorFastnessProvider;
        private IColorFastnessDetailProvider _IColorFastnessDetailProvider;

        public List<string> Get_Scales()
        {
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _IColorFastnessProvider.GetScales();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public FabricColorFastness_ViewModel Get_Main(string PoID)
        {
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            FabricColorFastness_ViewModel result = new FabricColorFastness_ViewModel();
            result.Result = true;
            try
            {
                if (string.IsNullOrEmpty(PoID))
                {
                    result.Result = false;
                    result.ErrorMessage = "PoID cannot be empty!";
                    return result;
                }

                var source = _IColorFastnessProvider.GetMain(PoID);
                if (string.IsNullOrEmpty(source.PoID ))
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                result = source;
                result.TargetLeadTime = _IColorFastnessProvider.Get_Target_LeadTime(result.CutInLine, result.MinSciDelivery);
                result.CompletionDate = (source.ArticlePercent >= 100) ? source.CompletionDate : null;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public ColorFastness_Result GetDetailHeader(string ID)
        {
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            ColorFastness_Result result = new ColorFastness_Result();
            try
            {
                if (string.IsNullOrEmpty(ID))
                {
                    result.baseResult.Result = false;
                    result.baseResult.ErrorMessage = "ID cannot be empty!";
                    return result;
                }

                var list = _IColorFastnessProvider.Get(ID);
                if (!list.Any() || list.Count() == 0)
                {
                    result.baseResult.Result = false;
                    result.baseResult.ErrorMessage = "Data not found!";
                    return result;
                }

                result = list.FirstOrDefault();
                result.baseResult = new DatabaseObject.BaseResult();
                result.baseResult.Result = true;
                result.baseResult.ErrorMessage = "";
            }
            catch (Exception ex)
            {
                result.baseResult.Result = false;
                result.baseResult.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public Fabric_ColorFastness_Detail_ViewModel GetDetailBody(string ID)
        {
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            Fabric_ColorFastness_Detail_ViewModel result = new Fabric_ColorFastness_Detail_ViewModel();
            try
            {
                result = _IColorFastnessDetailProvider.Get_DetailBody(ID);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return result;
        }

        public IList<PO_Supp_Detail> GetSeq(string POID, string Seq1 = "", string Seq2 = "")
        {
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _IColorFastnessDetailProvider.Get_Seq(POID, Seq1, Seq2);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<FtyInventory> GetRoll(string POID, string Seq1, string Seq2)
        {
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            try
            {
                return _IColorFastnessDetailProvider.Get_Roll(POID, Seq1, Seq2);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public BaseResult Save_ColorFastness_1stPage(string PoID, string Remark, List<ColorFastness_Result> _ColorFastness)
        {
            BaseResult baseResult = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IColorFastnessProvider = new ColorFastnessProvider(_ISQLDataTransaction);
                baseResult.Result = _IColorFastnessProvider.Save_PO(PoID, Remark);
                
                // 刪除前端傳來卻"不存在"DB的資料
                baseResult.Result = _IColorFastnessProvider.Delete_ColorFastness(PoID, _ColorFastness);
                _ISQLDataTransaction.Commit();
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return baseResult;
        }

        public BaseResult Save_ColorFastness_2ndPage(Fabric_ColorFastness_Detail_ViewModel source , string Mdivision, string UserID)
        {
            BaseResult baseResult = new BaseResult();
            SQLDataTransaction _ISQLDataTransaction = new SQLDataTransaction(Common.ProductionDataAccessLayer);
            try
            {
                _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(_ISQLDataTransaction);
                _IColorFastnessDetailProvider.Save_ColorFastness(source, Mdivision, UserID);                
                _ISQLDataTransaction.Commit();

                // 比對前端資料, 沒有的再刪除DB資料
                _IColorFastnessDetailProvider.Delete_ColorFastness_Detail(source.Main.ID, source.Detail);
                baseResult.Result = true;
            }
            catch (Exception ex)
            {
                _ISQLDataTransaction.RollBack();
                baseResult.Result = false;
                baseResult.ErrorMessage = ex.Message.ToString();
            }
            finally { _ISQLDataTransaction.CloseConnection(); }

            return baseResult;
        }
    }
}
