using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.SampleRFT;
using ProductionDataAccessLayer.Provider.MSSQL.StyleManagement;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessLogicLayer.Service.StyleManagement
{
    public class StyleResultService
    {
        public StyleResultProvider _Provider { get; set; }


        public StyleResult_ViewModel Get_StyleResult_Browse(StyleResult_Request styleResult_Request)
        {
            StyleResult_ViewModel result = new StyleResult_ViewModel()
            {
                SampleRFT = new List<StyleResult_SampleRFT>(),
                FTYDisclaimer = new List<StyleResult_FTYDisclaimer>(),
                RRLR = new List<StyleResult_RRLR>(),
                BulkFGT = new List<StyleResult_BulkFGT>(),
                PoList = new List<StyleResult_PoList>()
            };
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);

                styleResult_Request.CallType = StyleResult_Request.EnumCallType.StyleResult;

                bool HasSampleRFTInspection = _Provider.Check_SampleRFTInspection_Count(styleResult_Request) > 0 ? true : false;

                result.HasSampleRFTInspection = HasSampleRFTInspection;
                if (HasSampleRFTInspection)
                {
                    styleResult_Request.InspectionTableName = "SampleRFTInspection";
                }

                var styleResults = _Provider.Get_StyleInfo(styleResult_Request).ToList();

                if (styleResults.Any())
                {
                    result = styleResults.FirstOrDefault();
                }

                if (HasSampleRFTInspection)
                {
                    result.SampleRFT = _Provider.Get_StyleResult_SampleRFT_FromSampleRFTInspection(styleResult_Request).ToList();
                }
                else
                {
                    result.SampleRFT = _Provider.Get_StyleResult_SampleRFT(styleResult_Request).ToList();
                }

                result.FTYDisclaimer = _Provider.Get_StyleResult_FTYDisclaimer(styleResult_Request).ToList();
                result.RRLR = _Provider.Get_StyleResult_RRLR(styleResult_Request).ToList();
                result.BulkFGT = _Provider.Get_StyleResult_BulkFGT(styleResult_Request).ToList();
                result.PoList = _Provider.Get_StyleResult_PoList(styleResult_Request).ToList();


                result.Result = true;

            }
            catch (Exception ex)
            {
                result.Result = false;
                result.MsgScript = $@"msg.WithInfo('{ex.Message}');";
            }

            return result;
        }


        /// <summary>
        /// 透過  Microsoft.Office.Interop.Excel 生成Excel，並下載至暫存路徑
        /// </summary>
        /// <param name="Req"></param>
        /// <returns>暫存檔路徑</returns>
        public StyleResult_ViewModel GetExcel(StyleResult_Request Req)
        {
            string TempTilePath = string.Empty;
            StyleResult_ViewModel result = new StyleResult_ViewModel();
            DataTable dt = new DataTable();
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);

                bool HasSampleRFTInspection = _Provider.Check_SampleRFTInspection_Count(Req) > 0 ? true : false;

                if (HasSampleRFTInspection)
                {
                    Req.InspectionTableName = "SampleRFTInspection";
                }

                result = this.Get_StyleResult_Browse(Req);

                // 取得Datatable
                if (HasSampleRFTInspection)
                {
                    dt = _Provider.Get_StyleResult_SampleRFT_DataTable_FromSampleRFTInspection(Req);
                }
                else
                {
                    dt = _Provider.Get_StyleResult_SampleRFT_DataTable(Req);
                }


                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                }

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                }

                // 開啟excel app
                Excel.Application excelApp = MyUtility.Excel.ConnectExcel(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\StyleResult_SampleRFT.xltx");

                Excel.Worksheet worksheet = excelApp.Sheets[1];

                int RowIdx = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    worksheet.Cells[RowIdx + 2, 1] = dt.Rows[RowIdx]["SP"].ToString();
                    worksheet.Cells[RowIdx + 2, 2] = dt.Rows[RowIdx]["SampleStage"].ToString();
                    worksheet.Cells[RowIdx + 2, 3] = dt.Rows[RowIdx]["Factory"].ToString();
                    worksheet.Cells[RowIdx + 2, 4] = ((DateTime?)dt.Rows[RowIdx]["Delivery"]).HasValue ? ((DateTime?)dt.Rows[RowIdx]["Delivery"]).Value.ToShortDateString() : "";
                    worksheet.Cells[RowIdx + 2, 5] = ((DateTime?)dt.Rows[RowIdx]["SCIDelivery"]).HasValue ? ((DateTime?)dt.Rows[RowIdx]["SCIDelivery"]).Value.ToShortDateString() : ""; ;
                    worksheet.Cells[RowIdx + 2, 6] = dt.Rows[RowIdx]["InspectedQty"].ToString();
                    worksheet.Cells[RowIdx + 2, 7] = dt.Rows[RowIdx]["RFT"].ToString();
                    worksheet.Cells[RowIdx + 2, 8] = dt.Rows[RowIdx]["BAProduct"].ToString();
                    worksheet.Cells[RowIdx + 2, 9] = dt.Rows[RowIdx]["BAAuditCriteria"].ToString();
                    RowIdx++;
                }

                string fileName = $"Style Result _Sample RFT{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);

                result.TempFileName = fileName;
                result.Result = true;
            }
            catch (Exception ex)
            {
                result.Result = false;
                result.MsgScript = ex.Message;
            }

            return result;
        }

        public IList<StyleResult_Request> GetStyle(StyleResult_Request Req)
        {
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);
                return _Provider.GetStyle(Req).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public IList<SelectListItem> GetStyles(StyleResult_Request Req)
        {
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);
                return _Provider.GetStyles(Req).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<SelectListItem> GetBrands()
        {
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);
                return _Provider.GetBrands().ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IList<SelectListItem> GetSeasons(string brandID = "")
        {
            try
            {
                _Provider = new StyleResultProvider(Common.ProductionDataAccessLayer);
                return _Provider.GetSeasons(brandID).ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
