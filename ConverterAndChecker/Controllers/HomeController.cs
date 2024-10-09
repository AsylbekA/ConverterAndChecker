﻿using ConverterAndChecker.Models;
using ConverterAndChecker.Models.Excel;
using ConverterAndChecker.Models.Pdf;
using ConverterAndChecker.Services;
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
        string pdfPath = "path/to/your/file.pdf";
        var model = new UploadViewModel
        {
            PdfFile = CreateMockFile("C:\\perneke\\Пернеке ревизор документы\\мед страх\\2023 Апрель.pdf"),
            XlsxFile = CreateMockFile("C:\\perneke\\Пернеке ревизор документы\\Ведомость 2023\\2023 апрель опв.xls")
        };
        return Index(model);
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
        for (int i=0; i<2;i++)
        {
            var pdfText = _converer.ExtractTextFromPdf(model.PdfFile,i);
            foreach (var val in pdfText)
            {
                string key = val.IIN;
                if (pdfKeyValuePairs.ContainsKey(key))
                {
                    var temp = pdfKeyValuePairs[key];
                    temp.PdfTable.Add(val);
                    temp.FullInfo = temp.FullInfo + "\n" +  val.AccountNumber;
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


       // var stream = _converer.SetExcelInshurance(ExcelKeyValuePairs, pdfKeyValuePairs);
        var stream = _converer.SetExcelOPV(ExcelKeyValuePairs, pdfKeyValuePairs);
         
        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Результат по " + model.XlsxFile.Name + " НПФ" + ".xlsx");
        //var workbook = _converer.setExcelValue(model.XlsxFile, diffPdfExclSum);

        //byte[] modifiedWorkbookBytes = _converer.GetModifiedWorkbookBytes(workbook);

        //return File(modifiedWorkbookBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Результат по " + _converer.fileName);
    }
}
