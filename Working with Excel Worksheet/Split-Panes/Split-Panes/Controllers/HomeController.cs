using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Split_Panes.Models;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;

namespace Split_Panes.Controllers
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
        public IActionResult SplitPanes()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                #region Split Panes
                //split panes
                sheet.FirstVisibleColumn = 2;
                sheet.FirstVisibleRow = 5;
                sheet.VerticalSplit = 5000;
                sheet.HorizontalSplit = 5000;
                #endregion

                sheet.ActivePane = 1;

                //Saving the workbook
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "SplitPanes.xlsx";
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
