using Group_Ungroup_Rows_Columns.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using Syncfusion.XlsIO;
using System.IO;

namespace Group_Ungroup_Rows_Columns.Controllers
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
        public IActionResult GroupRowsColumns()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("InputTemplate - ToGroup.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Group Rows
                //Group Rows
                worksheet.Range["A3:A7"].Group(ExcelGroupBy.ByRows, true);
                worksheet.Range["A11:A16"].Group(ExcelGroupBy.ByRows);
                #endregion

                #region Group Columns
                //Group Columns
                worksheet.Range["C1:D1"].Group(ExcelGroupBy.ByColumns, false);
                worksheet.Range["F1:G1"].Group(ExcelGroupBy.ByColumns);
                #endregion

                //Saving the workbook
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "GroupRowsColumns.xlsx";
                return fileStreamResult;
            }
        }
        public IActionResult UngroupRowsColumns()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("InputTemplate - ToUngroup.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Un-Group Rows
                //Ungroup Rows
                worksheet.Range["A3:A7"].Ungroup(ExcelGroupBy.ByRows);
                #endregion

                #region Un-Group Columns
                //Ungroup Columns
                worksheet.Range["C1:D1"].Ungroup(ExcelGroupBy.ByColumns);
                #endregion

                //Saving the workbook
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "UngroupRowsColumns.xlsx";
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
