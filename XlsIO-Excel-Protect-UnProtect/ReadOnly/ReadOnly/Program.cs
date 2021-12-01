using Syncfusion.XlsIO;
using System;
using System.IO;

namespace ReadOnly
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                FileStream input = new FileStream("../../../Data/InputData.xlsx", FileMode.Open, FileAccess.ReadWrite);

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(input);

                //Set Read only
                workbook.ReadOnlyRecommended = true;

                FileStream output = new FileStream("../../../Output/ReadOnlyOutput.xlsx", FileMode.Create, FileAccess.ReadWrite);

                //Save the document
                workbook.SaveAs(output);
            }
        }
    }
}
