using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Save_HTML.Models;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;
using Syncfusion.XlsIO.Implementation;

namespace Save_HTML.Controllers
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
        public IActionResult SaveHTML()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Initialize excel engine and open workbook
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                worksheet.Range["A1:M20"].Text = "Html Document";

                //Create the instant for SaveOptions
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                saveOptions.TextMode = HtmlSaveOptions.GetText.DisplayText;

                //Saving the workbook
                MemoryStream stream = new MemoryStream();

                #region Save as HTML
                workbook.SaveAsHtml(stream, saveOptions);
                #endregion

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "text/html");
                fileStreamResult.FileDownloadName = "HTMLFile.html";
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
