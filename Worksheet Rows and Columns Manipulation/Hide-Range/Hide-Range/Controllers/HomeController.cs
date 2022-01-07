using Hide_Range.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;
using Syncfusion.XlsIO.Implementation.Collections;

namespace Hide_Range.Controllers
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
        public IActionResult HideRange()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                IRange range = worksheet.Range["D4"];

                #region Hide single cell
                //Hiding the range ‘D4’
                worksheet.ShowRange(range, false);
                #endregion

                IRange firstRange = worksheet.Range["F6:I9"];
                IRange secondRange = worksheet.Range["C15:G20"];
                RangesCollection rangeCollection = new RangesCollection(application, worksheet);
                rangeCollection.Add(firstRange);
                rangeCollection.Add(secondRange);

                #region Hide multiple cells
                //Hiding a collection of ranges
                worksheet.ShowRange(rangeCollection, false);
                #endregion

                //Saving the workbook
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "HideRange.xlsx";
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
