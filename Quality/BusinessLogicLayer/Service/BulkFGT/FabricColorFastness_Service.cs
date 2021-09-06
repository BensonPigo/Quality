using BusinessLogicLayer.Interface.BulkFGT;
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
            try
            {
                if (string.IsNullOrEmpty(PoID))
                {
                    result.Result = false;
                    result.ErrorMessage = "PoID cannot be empty!";
                    return result;
                }

                var list = _IColorFastnessProvider.GetMain(PoID);
                if (!list.Any() || list.Count() == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }

                result = list.FirstOrDefault();
                result.TargetLeadTime = _IColorFastnessProvider.Get_Target_LeadTime(result.EarliestDate, result.EarliestSCIDel);
                result.CompletionDate = (list.FirstOrDefault().ArticlePercent >= 100) ? list.FirstOrDefault().CompletionDate : null;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public FabricColorFastness_ViewModel GetDetailHeader(string ID)
        {
            _IColorFastnessProvider = new ColorFastnessProvider(Common.ProductionDataAccessLayer);
            FabricColorFastness_ViewModel result = new FabricColorFastness_ViewModel();
            try
            {
                if (string.IsNullOrEmpty(ID))
                {
                    result.Result = false;
                    result.ErrorMessage = "ID cannot be empty!";
                    return result;
                }
                var list = _IColorFastnessProvider.Get(ID);
                if (!list.Any() || list.Count() == 0)
                {
                    result.Result = false;
                    result.ErrorMessage = "Data not found!";
                    return result;
                }
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.ErrorMessage = ex.Message.ToString();
            }

            return result;
        }

        public IList<Fabirc_ColorFastness_Detail_ViewModel> GetDetailBody(string ID)
        {
            _IColorFastnessDetailProvider = new ColorFastnessDetailProvider(Common.ProductionDataAccessLayer);
            IList<Fabirc_ColorFastness_Detail_ViewModel> result = new List<Fabirc_ColorFastness_Detail_ViewModel>();
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
    }
}
