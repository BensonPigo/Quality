using DatabaseObject.ViewModel;
using System.Collections.Generic;
using ManufacturingExecutionDataAccessLayer.Interface;
using ManufacturingExecutionDataAccessLayer.Provider.MSSQL;
using System.Linq;
using BusinessLogicLayer.Interface.SampleRFT;
using System.IO;
using Ict;
using Sci.Data;
using System;
using System.Data;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Sci;

namespace BusinessLogicLayer.Service.SampleRFT
{
    public class CFTCommentsService : ICFTCommentsService
    {
        private ICFTCommentsProvider _CFTCommentsProvider;

        public CFTComments_ViewModel Get_CFT_Orders(CFTComments_ViewModel Req)
        {
            CFTComments_ViewModel model = new CFTComments_ViewModel();
            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ProductionDataAccessLayer);
                var res = _CFTCommentsProvider.Get_CFT_Orders(Req).ToList();

                if (res.Any())
                {
                    model = res.FirstOrDefault();
                }
                
                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        public CFTComments_ViewModel Get_CFT_OrderComments(CFTComments_ViewModel Req)
        {
            Req.DataList = new List<CFTComments_Result>();
            CFTComments_ViewModel model = Req;

            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);
                var res = _CFTCommentsProvider.Get_CFT_OrderComments(Req).ToList();

                model.DataList = res;
                model.Result = true;
            }
            catch (Exception ex)
            {
                model.Result = false;
                model.ErrorMessage = ex.Message;
            }

            return model;
        }

        /// <summary>
        /// 廢棄：生成Excel，並下載至暫存路徑
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public DataTable GetExcel_DataTable(CFTComments_ViewModel Req)
        {
            string TempTilePath = string.Empty;
            CFTComments_ViewModel result = new CFTComments_ViewModel();
            DataTable dt = new DataTable();
            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得Datatable
                dt = _CFTCommentsProvider.Get_CFT_OrderComments_DataTable(Req);

            }
            catch (Exception ex)
            {
                throw ex;
                //result.Result = false;
                //result.ErrorMessage = ex.Message;
            }

            return dt;
        }

        /// <summary>
        /// 廢棄：生成Excel，並下載至暫存路徑
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public CFTComments_ViewModel GetExcel(CFTComments_ViewModel Req)
        {
            string TempTilePath = string.Empty;
            CFTComments_ViewModel result = new CFTComments_ViewModel();
            DataTable dt = new DataTable();
            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得Datatable
                dt = _CFTCommentsProvider.Get_CFT_OrderComments_DataTable(Req);

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                }

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                }

                // 開啟excel app
                Excel.Application excelApp = MyUtility.Excel.ConnectExcel(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\CFT Comments.xltx");

                Excel.Worksheet worksheet = excelApp.Sheets[1];

                int RowIdx = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    worksheet.Cells[RowIdx + 2, 1] = dt.Rows[RowIdx]["SampleStage"].ToString();
                    worksheet.Cells[RowIdx + 2, 2] = dt.Rows[RowIdx]["CommentsCategory"].ToString();
                    worksheet.Cells[RowIdx + 2, 3] = dt.Rows[RowIdx]["Comnments"].ToString();
                    RowIdx++;
                }

                string fileName = $"CFT Comments{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);

                result.TempFilePath = filepath;
                result.Result = true;
            }
            catch (Exception ex)
            {
                throw ex;
                //result.Result = false;
                //result.ErrorMessage = ex.Message;
            }

            return result;
        }
    }
}
