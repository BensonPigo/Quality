using DatabaseObject.ViewModel.StyleManagement;
using ProductionDataAccessLayer.Provider.MSSQL.StyleManagement;
using Sci;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;

namespace BusinessLogicLayer.Service.StyleManagement
{
    public class ExceptionFDService
    {
        private ExceptionFDProvider _Provider;

        public List<ExceptionFD_ViewModel> GetData()
        {
            _Provider = new ExceptionFDProvider(Common.ProductionDataAccessLayer);
            List<ExceptionFD_ViewModel> result = _Provider.GetData().ToList();
            return result;
        }

        /// <summary>
        /// 透過  Microsoft.Office.Interop.Excel 生成Excel，並下載至暫存路徑
        /// </summary>
        /// <param name="Req"></param>
        /// <returns>暫存檔路徑</returns>
        public string GetExcel()
        {
            string TempTilePath = string.Empty;
            System.Data.DataTable dt = new System.Data.DataTable();
            Application excelApp = null;
            try
            {
                _Provider = new ExceptionFDProvider(Common.ProductionDataAccessLayer);

                // 取得Datatable
                dt = _Provider.GetDataTable();

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\");
                }

                if (!System.IO.Directory.Exists(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\"))
                {
                    System.IO.Directory.CreateDirectory(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\TMP\\");
                }

                // 開啟excel app
                excelApp = MyUtility.Excel.ConnectExcel(System.Web.HttpContext.Current.Server.MapPath("~/") + "\\XLT\\ExceptionFD.xltx");

                Worksheet worksheet = excelApp.Sheets[1];

                int RowIdx = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    worksheet.Cells[RowIdx + 2, 1] = dt.Rows[RowIdx]["StyleID"].ToString();
                    worksheet.Cells[RowIdx + 2, 2] = dt.Rows[RowIdx]["BrandID"].ToString();
                    worksheet.Cells[RowIdx + 2, 3] = dt.Rows[RowIdx]["SeasonID"].ToString();
                    worksheet.Cells[RowIdx + 2, 4] = dt.Rows[RowIdx]["Article"].ToString();
                    worksheet.Cells[RowIdx + 2, 5] = dt.Rows[RowIdx]["ExpectionFormStatus"].ToString();
                    worksheet.Cells[RowIdx + 2, 6] = dt.Rows[RowIdx]["ExpectionFormDate"];
                    worksheet.Cells[RowIdx + 2, 7] = dt.Rows[RowIdx]["ExpectionFormRemark"].ToString();
                    RowIdx++;
                }

                string fileName = $"Exception FD{DateTime.Now.ToString("yyyyMMdd")}{Guid.NewGuid()}.xlsx";
                string filepath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("~/"), "TMP", fileName);

                Workbook workbook = excelApp.ActiveWorkbook;
                workbook.SaveAs(filepath);

                workbook.Close();

                TempTilePath = fileName;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                MyUtility.Excel.KillExcelProcess(excelApp);
            }

            return TempTilePath;
        }
    }
}
