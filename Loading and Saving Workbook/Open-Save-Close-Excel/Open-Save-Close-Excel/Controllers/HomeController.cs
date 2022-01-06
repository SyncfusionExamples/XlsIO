using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Open_Save_Close_Excel.Models;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;

namespace Open_Save_Close_Excel.Controllers
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
        public IActionResult OpenSaveandCloseExcel()
        {
            //Creates a new instance for ExcelEngine
            ExcelEngine excelEngine = new ExcelEngine();

            #region Open
            //Loads or open an existing workbook through Open method of IWorkbook
            FileStream inputStream = new FileStream("InputTemplate.xlsx", FileMode.Open);
            IWorkbook workbook = excelEngine.Excel.Workbooks.Open(inputStream);
            #endregion

            //Set the version of the workbook
            workbook.Version = ExcelVersion.Xlsx;

            #region Save
            //Saving the workbook
            MemoryStream stream = new MemoryStream();
            workbook.SaveAs(stream);

            //Set the position as '0'.
            stream.Position = 0;

            //Download the Excel file in the browser
            FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
            fileStreamResult.FileDownloadName = "OpenandSave.xlsx";
            #endregion

            #region Close
            //Close the instance of IWorkbook
            workbook.Close();
            #endregion

            //Dispose the instance of ExcelEngine
            excelEngine.Dispose();
			
			return fileStreamResult;
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
