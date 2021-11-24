using Syncfusion.XlsIO;
using System;
using System.IO;

namespace ProtectWorkbook
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                FileStream input = new FileStream("../../../Data/InputWorkbook.xlsx", FileMode.Open, FileAccess.ReadWrite);

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(input);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Protect workbook with password
                workbook.Protect(true, true, "syncfusion");

                FileStream output = new FileStream("../../../Output/ProtectedWorkbook.xlsx", FileMode.Create, FileAccess.ReadWrite);

                //Save the modified document
                workbook.SaveAs(output);
            }
        }
    }
}
