using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Open_Save_CSV.Models;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;

namespace Open_Save_CSV.Controllers
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
        public IActionResult OpenandSaveCSV()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("InputTemplate.csv", FileMode.Open, FileAccess.Read);

                #region Open CSV
                //Open the Tab delimited CSV file
                IWorkbook workbook = application.Workbooks.Open(inputStream, "\t");
                #endregion

                IWorksheet worksheet = workbook.Worksheets[0];

                //Saving the workbook
                MemoryStream stream = new MemoryStream();

                #region Save CSV
                worksheet.SaveAs(stream, ",");
                #endregion

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "OpenandSaveCSV.csv";
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
