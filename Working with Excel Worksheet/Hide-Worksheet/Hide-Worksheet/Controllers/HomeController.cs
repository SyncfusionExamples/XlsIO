using Hide_Worksheet.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;

namespace Hide_Worksheet.Controllers
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
        public IActionResult HideWorksheet()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(2);
                IWorksheet sheet = workbook.Worksheets[0];

                sheet.Range["A1:M20"].Text = "visibility";

                #region Hide Worksheet
                //Set visibility
                sheet.Visibility = WorksheetVisibility.Hidden;
                #endregion

                //Saving the workbook
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "HideWorksheet.xlsx";
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
