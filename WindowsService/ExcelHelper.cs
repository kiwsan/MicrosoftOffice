using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService
{
    public static class ExcelHelper
    {

        public static void RunMacro(string filePath, params object[] parameters)
        {
            //Open the file and run the macro.
            Application oExcel = null;
            Workbooks oBooks = null;
            _Workbook oBook = null;
            object oMissing = Type.Missing;
            //Logger.Debug(string.Format(“Opening file { 0}to run macro”, filePath));
            try
            {
                oExcel = new Application();
                oExcel.Visible = false;
                oBooks = oExcel.Workbooks;

                oBook = oBooks.Open(filePath, oMissing, oMissing,
                    oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,
                    oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                // Run the macros.
                //RunMacro(oExcel, parameters);
                //Logger.Debug(string.Format(“Run macro successful for file { 0}”, filePath));
            }
            finally
            {
                if (oBook != null)
                {
                    oBook.Close(false, oMissing, oMissing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
                    oBook = null;
                }

                if (oBooks != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
                    oBooks = null;
                }

                if (oExcel != null)
                {
                    oExcel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
                    oExcel = null;
                }
            }
        }

    }
}
