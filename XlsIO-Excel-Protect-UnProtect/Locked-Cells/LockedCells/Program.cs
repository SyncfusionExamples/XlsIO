using System;
using System.IO;
using Syncfusion.XlsIO;

namespace LockedCells
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

                IWorksheet worksheet = workbook.Worksheets[0];

                //Unlock cell
                worksheet["A1"].CellStyle.Locked = false;

                FileStream output = new FileStream("../../../Output/LockedCells.xlsx", FileMode.Create, FileAccess.ReadWrite);

                //Save the modified document
                workbook.SaveAs(output);
            }
        }
    }
}
