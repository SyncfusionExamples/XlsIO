using Create_Access_Remove_Worksheet.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;

namespace Create_Access_Remove_Worksheet.Controllers
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
        public IActionResult CreateAccessRemove()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                #region Create
                //The new workbook is created with 5 worksheets
                IWorkbook workbook = application.Workbooks.Create(5);
                //Creating a new sheet
                IWorksheet worksheet = workbook.Worksheets.Create();
                //Creating a new sheet with name “Sample”
                IWorksheet namedSheet = workbook.Worksheets.Create("Sample");
                #endregion

                #region  Access
                //Accessing via index
                IWorksheet sheet = workbook.Worksheets[0];

                //Accessing via sheet name
                IWorksheet NamedSheet = workbook.Worksheets["Sample"];
                #endregion

                #region Remove
                //Removing the sheet
                workbook.Worksheets[0].Remove();
                #endregion

                //Saving the workbook
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "CreateAccessRemove.xlsx";
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
