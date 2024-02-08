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
            string xlsxData = ExtractDataFromXlsx(model.XlsxFile);

            return RedirectToAction("Index"); // Redirect to another action after uploading
        }


        public List<PaymentRecord> ExtractTextFromPdf(IFormFile pdfFile)
        {
            StringBuilder text = new StringBuilder();
            List<PaymentRecord> paymentRecords = new List<PaymentRecord>();
            using (PdfReader reader = new PdfReader(pdfFile.OpenReadStream()))
            {

                for (int page = 1; page <= reader.NumberOfPages; page++)
                {
                    string pageText = PdfTextExtractor.GetTextFromPage(reader, page);
                    ParsePageText(pageText, paymentRecords);
                }

            }

            return paymentRecords;
        }

        private void ParsePageText(string pageText, List<PaymentRecord> paymentRecords)
        {
            // Define regular expression patterns to match rows of the table and period information
            Regex rowPattern = new Regex(@"^\d+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+\S+\s+([\d,.]+)\s*$", RegexOptions.Multiline);
            Regex headerPattern = new Regex(@"№\s+п/п\s+Фамилия\s+Имя\s+Отчество\s+ИИН\s+Сумма\s*$");
            Regex periodPattern = new Regex(@"Период\s*:\s*([\d.]+)\s*-\s*([\d.]+)");

            // Find header, rows of the table, and period information
            Match headerMatch = headerPattern.Match(pageText);
            MatchCollection rowMatches = rowPattern.Matches(pageText);
            Match periodMatch = periodPattern.Match(pageText);

            // Determine the start and end index of the table
            int tableStartIndex = headerMatch.Index + headerMatch.Length;
            int tableEndIndex = rowMatches.Count > 0 ? rowMatches[rowMatches.Count - 1].Index + rowMatches[rowMatches.Count - 1].Length : pageText.Length;

            // Extract rows of the table
            string tableText = pageText.Substring(tableStartIndex, tableEndIndex - tableStartIndex);

            // Split the table text into rows
            string[] rows = tableText.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            // Extract period information
            string periodStart = periodMatch.Groups[1].Value;
            string periodEnd = periodMatch.Groups[2].Value;

            // Parse each row of the table
            foreach (string row in rows)
            {
                string[] rowData = row.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                // Extract data from the row and create a PaymentRecord object
                PaymentRecord record = new PaymentRecord
                {
                    LastName = rowData[2],
                    FirstName = rowData[3],
                    MiddleName = rowData[4],
                    IIN = rowData[5],
                    Amount = decimal.Parse(rowData[6].Replace(",", ""))
                };

                paymentRecords.Add(record);
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
