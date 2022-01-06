using Copy_Worksheet.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;

namespace Copy_Worksheet.Controllers
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
        public IActionResult CopyWorksheet()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream sourceStream = new FileStream("SourceTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook sourceWorkbook = application.Workbooks.Open(sourceStream);
                FileStream destinationStream = new FileStream("DestinationTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook destinationWorkbook = application.Workbooks.Open(destinationStream);

                #region Copy Worksheet
                //Copy first worksheet from the source workbook to the destination workbook
                destinationWorkbook.Worksheets.AddCopy(sourceWorkbook.Worksheets[0]);
                destinationWorkbook.ActiveSheetIndex = 1;
                #endregion

                //Saving the workbook
                MemoryStream stream = new MemoryStream();
                destinationWorkbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "CopyWorksheet.xlsx";
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
