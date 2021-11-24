using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Decrypt_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                application.DefaultVersion = ExcelVersion.Xlsx;

                FileStream input = new FileStream("../../../Data/EncryptedWorkbook.xlsx", FileMode.Open, FileAccess.ReadWrite);

                //Open encrypted Excel document with password
                IWorkbook workbook = application.Workbooks.Open(input, ExcelParseOptions.Default, false, "syncfusion");

                IWorksheet worksheet = workbook.Worksheets[0];

                //Decrypt workbook
                workbook.PasswordToOpen = string.Empty;

                FileStream output = new FileStream("../../../Output/DecryptedWorkbook.xlsx", FileMode.Create, FileAccess.ReadWrite);

                //Save the modified document
                workbook.SaveAs(output);
            }
        }
    }
}
