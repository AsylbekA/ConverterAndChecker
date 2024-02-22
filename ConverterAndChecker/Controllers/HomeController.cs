using ConverterAndChecker.Models;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;

namespace ConverterAndChecker.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    private DateTime StartDate { get; set; }
    private DateTime EndDate { get; set; }

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


        Dictionary<string, TableRows> keyValuePairs = new();

        foreach (var val in pdfText)
        {
            string key = val.Fio;
            if (keyValuePairs.ContainsKey(key))
            {
                var temp = keyValuePairs[key];
                temp.TableRow.Add(val);
                temp.Amount += val.Amount;
            }
            else
            {
                TableRows trs = new();
                trs.TableRow = new();
                trs.TableRow.Add(val);
                trs.Amount = val.Amount;
                keyValuePairs.Add(key, trs);
            }
        }



        ExtractDataFromExcel(model.XlsxFile);

        return RedirectToAction("Index"); // Redirect to another action after uploading
    }


    public List<TableRow> ParseTableFromPage(string pageText)
    {
        List<TableRow> rows = new List<TableRow>();

        // Define regular expression patterns to match each row of the table
        string pattern = @"(\d+)\s+(\p{L}+\s+\p{L}+\s+\p{L}+)\s+(\d{12})\s+(\S+)\s+([\d,]+\.\d{2})";
        Regex regex = new Regex(pattern, RegexOptions.Multiline);
        Regex periodPattern = new Regex(@"Период\s*:\s*([\d.]+)\s*-\s*([\d.]+)");


        // Match rows using the regular expression pattern
        MatchCollection matches = regex.Matches(pageText);
        Match periodMatch = periodPattern.Match(pageText);

        if (periodMatch.Captures.Count != 0)
        {
            StartDate = DateTime.Parse(periodMatch.Groups[1].Value);
            EndDate = DateTime.Parse(periodMatch.Groups[2].Value);
        }

        // Extract data from each match and create TableRow objects
        foreach (Match match in matches.Cast<Match>())
        {
            TableRow row = new();
            row.Number = match.Groups[1].Value;
            row.Fio = match.Groups[2].Value;
            row.IIN = match.Groups[3].Value;
            row.AccountNumber = match.Groups[4].Value;
            row.Amount = Convert.ToDecimal(match.Groups[5].Value.Replace(",", "").Replace(".", ","));
            rows.Add(row);
        }

        return rows;
    }
    public class TableRows
    {
        public Decimal Amount { get; set; }
        public List<TableRow> TableRow { get; set; }
    }

    public class TableRow
    {
        public string Number { get; set; }
        public string Fio { get; set; }
        public string IIN { get; set; }
        public string AccountNumber { get; set; }
        public decimal Amount { get; set; }
    }

    public List<TableRow> ExtractTextFromPdf(IFormFile pdfFile)
    {
        StringBuilder tableData = new StringBuilder();
        List<TableRow> pdfdatas = new();
        using (PdfReader reader = new PdfReader(pdfFile.OpenReadStream()))
        {
            for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber++)
            {
                string pageText = PdfTextExtractor.GetTextFromPage(reader, pageNumber);
                pdfdatas.AddRange(ParseTableFromPage(pageText));
            }
        }

        return pdfdatas;
    }

    //public void Ex

    public void ExtractDataFromExcel(IFormFile excelFile)
    {
        using var stream = excelFile.OpenReadStream();
        IWorkbook workbook;
        if (excelFile.FileName.EndsWith(".xlsx"))
            workbook = new XSSFWorkbook(stream);
        else if (excelFile.FileName.EndsWith(".xls"))   
            workbook = new HSSFWorkbook(stream);
        else
            throw new Exception("Unsupported file format.");

        ISheet sheet = workbook.GetSheetAt(0); // Assuming the data is in the first sheet

        // Loop through rows and columns to extract data
        for (int row = 0; row <= sheet.LastRowNum; row++)
        {
            IRow currentRow = sheet.GetRow(row);
            if (currentRow != null)
            {
                for (int col = 0; col < currentRow.Cells.Count; col++)
                {
                    Console.Write(currentRow.GetCell(col)?.ToString() + "\t");
                }
                Console.WriteLine();
            }
        }
    }
}
