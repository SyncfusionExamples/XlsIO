using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Mark_As_Final
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                FileStream input = new FileStream("../../../Data/InputData.xlsx", FileMode.Open, FileAccess.ReadWrite);

                IWorkbook workbook = application.Workbooks.Open(input);

                //Set workbook as final
                workbook.MarkAsFinal();

                FileStream output = new FileStream("../../../Output/MarkAsFinalOutput.xlsx", FileMode.Create, FileAccess.ReadWrite);

                //Save the modified document
                workbook.SaveAs(output);
            }
        }
    }
}
