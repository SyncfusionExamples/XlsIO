using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Resize_Rows_Columns.Models;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;

namespace Resize_Rows_Columns.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult ResizeRowsColumns()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Resize rows
                //Modifying the row height of one row
                worksheet.SetRowHeight(2, 100);
				
				//Modifying the row height of multiple rows
                worksheet.Range["A5:A10"].RowHeight = 40;                
                #endregion

                #region Resize columns
                //Modifying the column width of one column
                worksheet.SetColumnWidth(2, 50);

                //Modifying the column width of multiple columns
                worksheet.Range["D1:G1"].ColumnWidth = 5;
                #endregion

                //Saving the workbook
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "ResizeRowsColumns.xlsx";
                return fileStreamResult;
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
