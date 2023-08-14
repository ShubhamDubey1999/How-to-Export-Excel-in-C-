using Dataset_To_XLSX_Format.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Diagnostics;
using ClosedXML.Excel;
using SautinSoft.Document;

namespace Dataset_To_XLSX_Format.Controllers
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
        public IActionResult Index1()
        {
            return View();
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

        public IActionResult Report()
        {
            return View();
        }

        private DataSet GetDataSet()
        {
            DataSet ds = new DataSet();
            DataTable dtbl = new DataTable();
            dtbl.Columns.Add("Sl No");//Define Columns
            dtbl.Columns.Add("Novel Name");
            dtbl.Columns.Add("Author");
            dtbl.Columns.Add("Genres");
            dtbl.Columns.Add("Published Date");
            dtbl.Columns.Add("Price");
            dtbl.Columns.Add("Rating");

            dtbl.Rows.Add("1", "In Search of Lost Time", "Marcel Proust", "Literary modernism", "01-01-1913", "348", "4.3");//Adding Rows
            dtbl.Rows.Add("2", "Ulysses", "James Joyce", "Modernism", "22-02-1922", "58", "3.7");
            dtbl.Rows.Add("3", "Moby Dick", "Herman Melville", "Adventure fiction", "18-10-1851", "131", "3.4");
            dtbl.Rows.Add("4", "Hamlet", "William Shakespeare", "Tragedy", "01-01-1603", "225", "3.9");
            dtbl.Rows.Add("5", "War and Peace", "Leo Tolstoy", "Historical fiction", "01-01-1869", "133.95", "4.1");
            dtbl.TableName = "Sheet1";
            ds.Tables.Add(dtbl);

            DataTable dtbl2 = dtbl.Copy();//Created copies of first table
            dtbl2.TableName = "Sheet2";
            dtbl2.Rows.Add("6", "War and Peace_V2", "Leo Tolstoy_V2", "Historical fiction_V2", "01-01-1869_V2", "133.95_V2", "4.1_V2");
            ds.Tables.Add(dtbl2);
            DataTable dtbl3 = dtbl.Copy();//Created copies of first table
            dtbl3.TableName = "Sheet3";
            dtbl3.Rows.Add("6", "War and Peace_V3", "Leo Tolstoy_V3", "Historical fiction_V3", "01-01-1869_V3", "133.95_V3", "4.1_V3");
            ds.Tables.Add(dtbl3);
            return ds;
        }

        public IActionResult DownloadFile_Excel()
        {
            XLWorkbook wb = new XLWorkbook();
            var ds = GetDataSet();

            //Add Hyperlink Sheet
            var ws = wb.Worksheets.Add("Hyperlinks");
            for (var i = 1; i <= ds.Tables.Count; i++)
            {
                var Sheet_Name = "Sheet" + i.ToString();
                ws.Cell(i, 1).Value = Sheet_Name;
                ws.Cell(i, 1).SetHyperlink(new XLHyperlink("'" + Sheet_Name + "'!A1"));
            }

            //Adding Rest of the sheets
            foreach (DataTable dt in ds.Tables)
            {
                var worksheet = wb.Worksheets.Add(dt);
                worksheet.AutoFilter.Clear();
                var r = dt.Rows.Count;
                var c = dt.Columns.Count;
                var lstCol = (char)(64 + c);
                var range = "A1:" + lstCol + c.ToString();
                worksheet.Range(range).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Range(range).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Range(range).Style.Font.Italic = true;

                worksheet.Range("A1:G1").Style.Font.SetFontSize(20.0d);
                worksheet.Range("A1:G1").Style.Fill.SetBackgroundColor(XLColor.Brown);
            }

            using (MemoryStream stream = new MemoryStream())
            {
                wb.SaveAs(stream);
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Grid.xlsx");
            }
        }

        public IActionResult DownloadFile_Word()
        {
            // cshtml File Path
            var inpFile1 = Path.Combine(Directory.GetCurrentDirectory(), "Views", "Home", "Index1.cshtml");
            // html File Path
            var inpFile2 = "C:\\Users\\rmani\\ABHAY Programs & Notes\\MY PROGRAMS\\frontend & backend development\\WebPage1\\html\\test.html";
            if (System.IO.File.Exists(inpFile2))
            {
                DocumentCore dc = DocumentCore.Load(inpFile2);
                using (MemoryStream stream = new MemoryStream())
                {
                    // Convert to PDF Format
                    //dc.Save(stream, new PdfSaveOptions());
                    //return File(stream.ToArray(), "application/pdf", "Result.pdf");

                    //Convert to Docx Format
                    dc.Save(stream, new DocxSaveOptions());
                    return File(stream.ToArray(), "application/msword", "Result.docx");
                }
            }
            else
                return Json("Fail");
        }



    }
}