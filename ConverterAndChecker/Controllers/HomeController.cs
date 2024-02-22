using ConverterAndChecker.Models;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
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
    private readonly IConfiguration _configuration;

    private readonly int cellCount;

    public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
    {
        _configuration = configuration;
        _logger = logger;
        cellCount = 53;
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

        var excelRow = ExtractDataFromExcel(model.XlsxFile);

        Dictionary<string, ExcelRows> ExcelKeyValuePairs = new();
        foreach (var val in excelRow)
        {
            string key = val.Fio;
            if (ExcelKeyValuePairs.ContainsKey(key))
            {
                var temp = ExcelKeyValuePairs[key];
                temp.ExcelRow.Add(val);
                temp.Amount += val.Amount;
            }
            else
            {
                ExcelRows ers = new();
                ers.ExcelRow = new();
                ers.ExcelRow.Add(val);
                ers.Amount = val.Amount;
                ExcelKeyValuePairs.Add(key, ers);
            }
        }


        Dictionary<string, (decimal, string, string)> diffPdfExclSum = new Dictionary<string, (decimal, string, string)>();
        string comment = "";
        string color = "green";
        foreach (var val in ExcelKeyValuePairs)
        {
            if (keyValuePairs.ContainsKey(val.Key))
            {
                var temp = keyValuePairs[val.Key];

                if (!diffPdfExclSum.ContainsKey(val.Key))
                {
                    decimal diffSum = val.Value.Amount - temp.Amount;

                    if (diffSum > 0)
                    {
                        comment = "Excel Amount more than PDF.  Profit = " + diffSum + "; ExcelSum = " + val.Value.Amount + " PdfSum = " + temp.Amount;
                        color = "yellow";
                    }
                    else if (diffSum < 0)
                    {
                        comment = "Pdf Amount more than Excel.  Deficit = " + diffSum + "; ExcelSum = " + val.Value.Amount + " PdfSum = " + temp.Amount;
                        color = "red";
                    }
                    diffPdfExclSum.Add(val.Key, (diffSum, comment, color));
                }
                keyValuePairs.Remove(val.Key);
            }
            else
            {
                diffPdfExclSum.Add(val.Key, (val.Value.Amount, "Not Found Client  from PDF", "rrrr"));
            }
        }

        foreach (var val in keyValuePairs)
        {
            diffPdfExclSum.Add(val.Key, (val.Value.Amount, "Not Found Client  from Excel", "ttt"));
        }


        return View(diffPdfExclSum);
        //return RedirectToAction("Index"); // Redirect to another action after uploading
    }
    public void saveExcel(IFormFile XlsxFile)
    {

        using var stream = XlsxFile.OpenReadStream();
        IWorkbook workbook;
        if (XlsxFile.FileName.EndsWith(".xlsx"))
            workbook = new XSSFWorkbook(stream);
        else if (XlsxFile.FileName.EndsWith(".xls"))
            workbook = new HSSFWorkbook(stream);
        else
            throw new Exception("Unsupported file format.");

        ISheet sheet = workbook.GetSheetAt(0);

        FileInfo inputFile = new FileInfo(XlsxFile.FileName);
        //using (ExcelPackage package = new ExcelPackage(inputFile))
        //{
        //    // Get the first worksheet in the Excel file
        //    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

        //    // Modify data as needed
        //    worksheet.Cells["A1"].Value = "Modified Value";

        //    // Save changes to a new Excel file
        //    FileInfo outputFile = new FileInfo(outputFilePath);
        //    package.SaveAs(outputFile);
        //}
    }

    public List<TableRow> ParseTableFromPage(string pageText)
    {
        List<TableRow> rows = new List<TableRow>();

        // Define regular expression patterns to match each row of the table
        string pattern = @"(\d+)\s+(\p{L}+\s+\p{L}+\s+\p{L}+)\s+(\d{12})\s+(\S+)\s+([\d,]+\.\d{2})";
        Regex regex = new(pattern, RegexOptions.Multiline);
        Regex periodPattern = new(@"Период\s*:\s*([\d.]+)\s*-\s*([\d.]+)");


        // Match rows using the regular expression pattern
        MatchCollection matches = regex.Matches(pageText);
        System.Text.RegularExpressions.Match periodMatch = periodPattern.Match(pageText);

        if (periodMatch.Captures.Count != 0)
        {
            StartDate = DateTime.Parse(periodMatch.Groups[1].Value);
            EndDate = DateTime.Parse(periodMatch.Groups[2].Value);
        }

        // Extract data from each match and create TableRow objects
        foreach (System.Text.RegularExpressions.Match match in matches.Cast<System.Text.RegularExpressions.Match>())
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

    public class ExcelRows
    {
        public Decimal Amount { get; set; }
        public List<ExcelRow> ExcelRow { get; set; }
    }

    public class TableRow
    {
        public string Number { get; set; }
        public string Fio { get; set; }
        public string IIN { get; set; }
        public string AccountNumber { get; set; }
        public decimal Amount { get; set; }
    }

    public class ExcelRow
    {
        public string Number { get; set; }
        public string Fio { get; set; }
        public decimal Amount { get; set; }
    }

    public List<TableRow> ExtractTextFromPdf(IFormFile pdfFile)
    {
        List<TableRow> pdfdatas = new();
        using (PdfReader reader = new(pdfFile.OpenReadStream()))
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

    public List<ExcelRow> ExtractDataFromExcel(IFormFile excelFile)
    {
        using var stream = excelFile.OpenReadStream();
        IWorkbook workbook;
        if (excelFile.FileName.EndsWith(".xlsx"))
            workbook = new XSSFWorkbook(stream);
        else if (excelFile.FileName.EndsWith(".xls"))
            workbook = new HSSFWorkbook(stream);
        else
            throw new Exception("Unsupported file format.");

        ISheet sheet = workbook.GetSheetAt(0);

        int transferredToTheBankCell = -1;
        int queueRow = -1;
        string queueIcon = "№ п/п";
        int fioRow = -1;
        bool hasHeader = false;
        List<ExcelRow> rows = new List<ExcelRow>();


        for (int row = 0; row <= sheet.LastRowNum; row++)
        {
            IRow currentRow = sheet.GetRow(row);
            if (currentRow != null)
            {
                if (!hasHeader)
                {
                    for (int col = 0; col < currentRow.Cells.Count; col++)
                    {
                        if (transferredToTheBankCell == -1 || queueRow == -1 || fioRow == -1)
                        {
                            string val = currentRow.GetCell(col)?.ToString() ?? "";
                            if (val.Equals("Перечислено в банк"))
                            {
                                transferredToTheBankCell = col;
                            }
                            else if (val.Equals(queueIcon))
                            {
                                queueRow = col;
                            }
                            else if (val.Equals("Физ. лицо"))
                            {
                                fioRow = col;
                            }
                        }
                        else
                        {
                            hasHeader = true;
                            break;
                        }
                        Console.Write(currentRow.GetCell(col)?.ToString() + "\t");
                    }
                }
                else
                {
                    ExcelRow excelRow = new ExcelRow();
                    string q = currentRow.GetCell(queueRow)?.ToString() ?? "";
                    int qq = -1;

                    try
                    {

                        qq = Convert.ToInt32(q);
                        if (qq != -1)
                        {
                            Console.WriteLine(qq);

                            excelRow.Number = q;
                            excelRow.Fio = currentRow.GetCell(fioRow)?.ToString() ?? "";
                            string sum = currentRow.GetCell(transferredToTheBankCell)?.ToString() ?? "0";
                            excelRow.Amount = Convert.ToDecimal(sum);
                            rows.Add(excelRow);
                        }
                    }
                    catch (Exception ex)
                    {
                    }

                }

                Console.WriteLine();
            }
        }

        return rows;
    }
}
