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
        /// 生成Excel，並下載至暫存路徑
        /// </summary>
        /// <param name="Req"></param>
        /// <returns></returns>
        public string GetExcel(CFTComments_ViewModel Req)
        {
            string TempTilePath = string.Empty;
            DataTable dt = new DataTable();
            try
            {
                _CFTCommentsProvider = new CFTCommentsProvider(Common.ManufacturingExecutionDataAccessLayer);

                // 取得Datatable
                dt = _CFTCommentsProvider.Get_CFT_OrderComments_DataTable(Req);

                // 開啟excel app
                Excel.Application excelApp = MyUtility.Excel.ConnectExcel(AppDomain.CurrentDomain.BaseDirectory + "XLT\\CFT Comments.xltx");

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
                string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TMP", fileName);

                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();
                excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);

                TempTilePath = filepath;
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
