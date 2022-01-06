using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Send_to_ClientBrowser.Models;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;

namespace Send_to_ClientBrowser.Controllers
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
        public IActionResult SendtoClientBrowser()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Loads or open an existing workbook through Open method of IWorkbook
                FileStream inputStream = new FileStream("InputTemplate.xlsx", FileMode.Open);
                IWorkbook workbook = excelEngine.Excel.Workbooks.Open(inputStream);

                //To-Do some manipulation
                //To-Do some manipulation

                //Initialize content type
                string ContentType = null;

                //Set the version of the workbook
                workbook.Version = ExcelVersion.Xlsx;
                ContentType = "Application/msexcel";

                //Save the workbook to stream
                MemoryStream outputStream = new MemoryStream();
                workbook.SaveAs(outputStream);
                outputStream.Position = 0;

                //Return the file with content type
                return File(outputStream, ContentType, "SendtoClientBrowser.xlsx");
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
