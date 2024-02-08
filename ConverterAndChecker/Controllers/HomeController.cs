using ConverterAndChecker.Models;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;

namespace ConverterAndChecker.Controllers
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

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


        // GET: /Home/Upload
        public IActionResult Upload()
        {
            return View();
        }

        // POST: /Home/Upload
        [HttpPost]
        public IActionResult Upload(UploadViewModel model)
        {
            // Handle file uploads here
            // Access uploaded files through model.PdfFile and model.XlsxFile
            var pdfText = ExtractTextFromPdf(model.PdfFile);
            string pdfText = ExtractTextFromPdf(model.PdfFile);
            string xlsxData = ExtractDataFromXlsx(model.XlsxFile);

            return RedirectToAction("Index"); // Redirect to another action after uploading
        }


        public string ExtractTextFromPdf(IFormFile pdfFile)
        {
            using (PdfReader reader = new PdfReader(pdfFile.OpenReadStream()))
            {
                StringBuilder text = new StringBuilder();
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
                return text.ToString();
            }

            // Output period information (for demonstration)
            Console.WriteLine($"Period: {periodStart} - {periodEnd}");
        }

        public string ExtractDataFromXlsx(IFormFile xlsxFile)
        {
            using (ExcelPackage package = new ExcelPackage(xlsxFile.OpenReadStream()))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet
                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;
                StringBuilder data = new StringBuilder();
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= columnCount; col++)
                    {
                        data.Append(worksheet.Cells[row, col].Value?.ToString() ?? "");
                    }
                }
                return data.ToString();
            }
        }
    }
}
