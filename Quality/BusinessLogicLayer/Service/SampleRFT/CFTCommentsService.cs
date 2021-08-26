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


        public string ToExcel(CFTComments_ViewModel Req)
        {
            string TempTilePath = string.Empty;
            DataTable dt;
            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得Datatable
                dt = _CFTCommentsProvider.Get_CFT_OrderComments_DataTable(Req);

                Excel.Application excelApp = MyUtility.Excel.ConnectExcel(AppDomain.CurrentDomain.BaseDirectory + "XLT\\CFT Comments.xltx");


                string xltx_name = "CFT Comments.xltx";

                MyUtility.Excel.CopyToXls(dt, string.Empty, xltx_name, 2, false, null, excelApp, wSheet: excelApp.Sheets[1]);



                #region 存檔 > 讀取MemoryStream > 下載 > 刪除
                string fileName = $"CFT Comments{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileName);
                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);
                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);

                TempTilePath = filepath;

                #endregion
            }
            catch (Exception ex)
            {
                //model.Result = false;
                //model.ErrorMessage = ex.Message;
            }

            return TempTilePath;
        }
    }
}
