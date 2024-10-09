using ConverterAndChecker.Models;
using ConverterAndChecker.Models.Excel;
using ConverterAndChecker.Models.Pdf;
using ConverterAndChecker.Services;
using iText.Kernel.Geom;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.AspNetCore.Mvc;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
namespace ConverterAndChecker.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;
    private readonly ConverterService _converer;

    public HomeController(ILogger<HomeController> logger, ConverterService converter)
    {
        _logger = logger;
        _converer = converter;
    }
    public IActionResult Index()
    {

        //string pdf = "2023 Март" + ".pdf";
        //string xmlP = "2023 март.xlsx";

        //string pdfPAth = "G:\\Work\\PernebekTaga\\04-10-2024\\Пернеке ревизор документы\\мед страх\\2023\\" + pdf;
        //string xlsPath = "\"G:\\Work\\PernebekTaga\\04-10-2024\\Пернеке ревизор документы\\Ведомость 2023\\2023 март.xlsx";// + xmlP;
        //var model = new UploadViewModel
        //{
            
        //    PdfFile = CreateMockFile(pdfPAth),
        //    XlsxFile = CreateMockFile(xlsPath)
        //};
        //return Index(model);

        return View();
    }
    

    private IFormFile CreateMockFile(string filePath)
    {
        var fileName = System.IO.Path.GetFileName(filePath);
        var memoryStream = new MemoryStream(System.IO.File.ReadAllBytes(filePath));
        var formFile = new FormFile(memoryStream, 0, memoryStream.Length, "file", fileName)
        {
            Headers = new HeaderDictionary(),
            ContentType = "application/octet-stream"
        };
        return formFile;
    }

    private static string[] itemsToCheck = new string[]
    {"16.08.2023 2523519/23-3446", "22.08.2023 2523519/23-3498", "22.08.2023 2523519/23-3459", "23.08.2023 2523519/23-3509", "23.08.2023 2523519/23-3311", "24.08.2023 2523519/23-3474", "24.08.2023 2523519/23-3473", "25.08.2023 2523519/23-3548", "25.08.2023 2523519/23-3727", "29.08.2023 2523519/23-3868" };

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
    // POST: /Home/Upload
    [HttpPost]
    public IActionResult Index(UploadViewModel model)
    {
        Dictionary<string, PdfTables> pdfKeyValuePairs = new();
       // List<string> itemsToCheck = new List<string>();
                var pdfText = _converer.ExtractTextFromPdf(model.PdfFile);


        //itemsToCheck = "16.08.2023 2523519/23-3446", "22.08.2023 2523519/23-3498", "22.08.2023 2523519/23-3459", "23.08.2023 2523519/23-3509", "23.08.2023 2523519/23-3311", "24.08.2023 2523519/23-3474", "24.08.2023 2523519/23-3473", "25.08.2023 2523519/23-3548", "25.08.2023 2523519/23-3727", "29.08.2023 2523519/23-3868"
            foreach (var val in pdfText)
            {

                bool containsAny = itemsToCheck.Any(item => val.ShortInfo.Contains(item));
                

                if (containsAny)
                {
                if (val.IIN == "780820402389")
                {
                    Console.WriteLine("scdscds");
                }
                    string key = val.IIN;
                    if (pdfKeyValuePairs.ContainsKey(key))
                    {
                        var temp = pdfKeyValuePairs[key];
                        temp.PdfTable.Add(val);
                        temp.FullInfo = temp.FullInfo + "\n" + val.AccountNumber;
                        temp.Amount += val.Amount;
                    }
                    else
                    {
                        PdfTables trs = new();
                        trs.PdfTable = new();
                        trs.PdfTable.Add(val);
                        trs.Amount = val.Amount;
                        trs.FullInfo = val.AccountNumber;
                        pdfKeyValuePairs.Add(key, trs);
                    }
                }
            }
       


        var excelRow = _converer.ExtractInshuranceFromExcel(model.XlsxFile);

        Dictionary<string, ExcelRows> ExcelKeyValuePairs = new();
        foreach (var val in excelRow)
        {
            string key = val.Iin;
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


        var stream = _converer.SetExcelInshurance(ExcelKeyValuePairs, pdfKeyValuePairs);
        //var stream = _converer.SetExcelOPV(ExcelKeyValuePairs, pdfKeyValuePairs);
        var name = model.XlsxFile.FileName.Replace(".xls", "");

        //name = model.XlsxFile.FileName.Replace(".xlsx", "");
        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Результат по " + name + " Мед страх" + ".xlsx");
        //var workbook = _converer.setExcelValue(model.XlsxFile, diffPdfExclSum);

        //byte[] modifiedWorkbookBytes = _converer.GetModifiedWorkbookBytes(workbook);

        //return File(modifiedWorkbookBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Результат по " + _converer.fileName);
    }
}
