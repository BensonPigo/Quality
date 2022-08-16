using Sci;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Library
{
    public static partial class ConvertToPDF
    {
        /// <summary>
        /// ExcelToPDF
        /// </summary>
        /// <param name="excelPath">ExcelPath</param>
        /// <param name="pdfPath">PDFPath</param>
        /// <returns>bool</returns>
        public static bool ExcelToPDF(string excelPath, string pdfPath)
        {
            bool result = false;
            Microsoft.Office.Interop.Excel.XlFixedFormatType targetType = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
            object missing = Type.Missing;
            Microsoft.Office.Interop.Excel.Application application = null;
            Microsoft.Office.Interop.Excel.Workbook workBook = null;
            try
            {
                application = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = false,
                };
                workBook = application.Workbooks.Open(excelPath);
                workBook.ExportAsFixedFormat(targetType, pdfPath);
                result = true;
            }
            catch (Exception)
            {
                result = false;
            }
            finally
            {
                if (application != null)
                {
                    MyUtility.Excel.KillExcelProcess(application);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return result;
        }

    }
}
