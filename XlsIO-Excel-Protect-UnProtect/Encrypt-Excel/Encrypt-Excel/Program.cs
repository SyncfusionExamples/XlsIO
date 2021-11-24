using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Encrypt_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                FileStream input = new FileStream("../../../Data/InputExcel.xlsx", FileMode.Open, FileAccess.ReadWrite);

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(input);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Encrypt workbook with password
                workbook.PasswordToOpen = "syncfusion";

                FileStream output = new FileStream("../../../Output/EncryptedWorkbook.xlsx", FileMode.Create, FileAccess.ReadWrite);

                //Save the modified document
                workbook.SaveAs(output);
            }
        }
    }
}
