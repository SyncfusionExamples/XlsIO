using Syncfusion.XlsIO;
using System;
using System.IO;

namespace UnProtectWorksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                FileStream input = new FileStream("../../../Data/ProtectedWorksheet.xlsx", FileMode.Open, FileAccess.ReadWrite);

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(input);

                IWorksheet worksheet = workbook.Worksheets[0];

                //UnProtect worksheet with password
                worksheet.Unprotect("syncfusion");

                FileStream output = new FileStream("../../../Output/UnProtectedSheet.xlsx", FileMode.Create, FileAccess.ReadWrite);

                //Save the modified document
                workbook.SaveAs(output);
            }
        }
    }
}
