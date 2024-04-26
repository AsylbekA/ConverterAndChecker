using ConverterAndChecker.Models;
using ConverterAndChecker.Models.Excel;
using ConverterAndChecker.Models.Pdf;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace ConverterAndChecker.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;
    private string fileName;
    private DateTime StartDate;
    private DateTime EndDate;

    string desimalString = "";
    decimal decimalDecimal;
    string firstReplase;
    string secondReplase;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }
    public IActionResult Index()
    {
        return View();
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
        var pdfText = ExtractTextFromPdf(model.PdfFile);
        Dictionary<string, PdfTables> pdfKeyValuePairs = new();
        foreach (var val in pdfText)
        {
            string key = val.Fio;
            if (pdfKeyValuePairs.ContainsKey(key))
            {
                var temp = pdfKeyValuePairs[key];
                temp.PdfTable.Add(val);
                temp.Amount += val.Amount;
            }
            else
            {
                PdfTables trs = new();
                trs.PdfTable = new();
                trs.PdfTable.Add(val);
                trs.Amount = val.Amount;
                pdfKeyValuePairs.Add(key, trs);
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

        foreach (var val in ExcelKeyValuePairs)
        {
            string upperkey = val.Key.ToUpper();
            if (pdfKeyValuePairs.ContainsKey(upperkey))
            {
                var temp = pdfKeyValuePairs[upperkey];

                if (!diffPdfExclSum.ContainsKey(upperkey))
                {
                    decimal diffSum = val.Value.Amount - temp.Amount;

                    string color;
                    string comment;
                    if (diffSum > 0)
                    {
                        comment = "Сумма из Расчетная ведомоста повышает сумму из Выписка, по проведенным платежам.  Розница повышение = " + diffSum + "; Сумма из Расчетная ведомоста = " + val.Value.Amount + "; Сумма из Выписка, по проведенным платежам = " + temp.Amount;
                        color = "yellow";
                    }
                    else if (diffSum < 0)
                    {
                        comment = "Сумма из Выписка, по проведенным платежам. нехватает из сумму  Расчетная ведомоста.  Недостаточная сумма = " + diffSum + "; Сумма из Расчетная ведомоста  = " + val.Value.Amount + "; Сумма из Выписка, по проведенным платежам = " + temp.Amount;
                        color = "darkyellow";
                    }
                    else
                    {
                        comment = "";
                        color = "green";
                    }
                    diffPdfExclSum.Add(val.Key, (diffSum, comment, color));
                }
            }
            else
            {
                diffPdfExclSum.Add(val.Key, (val.Value.Amount, "Не найден клиент из Выписка, по проведенным платежам", "red"));
            }

           pdfKeyValuePairs.Remove(upperkey);
        }

        foreach (var val in pdfKeyValuePairs)
        {
            diffPdfExclSum.Add(val.Key, (val.Value.Amount, "Не найден клиент из Расчетная ведомоста", "white"));
        }

        var workbook = setExcelValue(model.XlsxFile, diffPdfExclSum);

        byte[] modifiedWorkbookBytes = GetModifiedWorkbookBytes(workbook);

        return File(modifiedWorkbookBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Результат по " + fileName);
    }
    public byte[] GetModifiedWorkbookBytes(IWorkbook workbook)
    {
        // Create a memory stream to store the workbook data
        using (MemoryStream memoryStream = new MemoryStream())
        {
            // Write the workbook data to the memory stream
            workbook.Write(memoryStream);

            // Convert the memory stream to a byte array
            byte[] workbookBytes = memoryStream.ToArray();

            return workbookBytes;
        }
    }
    public IWorkbook setExcelValue(IFormFile excelFile, Dictionary<string, (decimal, string, string)> diffPdfExclSum)
    {
        using var stream = excelFile.OpenReadStream();
        IWorkbook workbook;
        if (excelFile.FileName.EndsWith(".xlsx"))
            workbook = new XSSFWorkbook(stream);
        else if (excelFile.FileName.EndsWith(".xls"))
            workbook = new HSSFWorkbook(stream);
        else
            throw new Exception("Unsupported file format.");


        fileName = excelFile.FileName;
        ISheet sheet = workbook.GetSheetAt(0);

        int transferredToTheBankCell = -1;
        int queueRowCell = -1;
        string queueIcon = "№ п/п";
        int fioRowCell = -1;
        bool hasHeader = false;
        int maxCellsCount = 0;

        for (int row = 0; row <= sheet.LastRowNum; row++)
        {
            IRow currentRow = sheet.GetRow(row);
            if (currentRow != null)
            {
                if (!hasHeader)
                {
                    if (maxCellsCount < currentRow.Cells.Count) maxCellsCount = currentRow.Cells.Count;
                    for (int col = 0; col < currentRow.Cells.Count; col++)
                    {
                        if (transferredToTheBankCell == -1 || queueRowCell == -1 || fioRowCell == -1)
                        {
                            string val = currentRow.GetCell(col)?.ToString() ?? "";
                            if (val.Equals("Перечислено в банк"))
                            {
                                transferredToTheBankCell = col;
                            }
                            else if (val.Equals(queueIcon))
                            {
                                queueRowCell = col;
                            }
                            else if (val.Equals("Физ. лицо"))
                            {
                                fioRowCell = col;
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
                    string queueCell = currentRow.GetCell(queueRowCell)?.ToString() ?? "";
                    int convertedQueueCell = -1;

                    try
                    {
                        convertedQueueCell = Convert.ToInt32(queueCell);
                        if (convertedQueueCell != -1)
                        {
                            Console.WriteLine(convertedQueueCell);
                            string fioKey = currentRow.GetCell(fioRowCell)?.ToString() ?? "";
                            if (diffPdfExclSum.ContainsKey(fioKey))
                            {
                                var clientValue = diffPdfExclSum[fioKey];

                                // Set the value of cell A1 to "New Value"

                                NPOI.SS.UserModel.ICell newCell = currentRow.GetCell(maxCellsCount + 2) ?? currentRow.CreateCell(maxCellsCount + 2); // Get the first cell or create a new one if it doesn't exist
                                // Create a new cell style
                                if (clientValue.Item3.Equals("green"))
                                {
                                    ICellStyle cellStyle = workbook.CreateCellStyle();
                                    // Set the fill foreground color (you can change this to any color you desire)
                                    cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index; // for example, light orange

                                    // Set the fill pattern
                                    cellStyle.FillPattern = FillPattern.SolidForeground;

                                    // Apply the style to the cell
                                    newCell.CellStyle = cellStyle;
                                }
                                else if (clientValue.Item3.Equals("yellow"))
                                {
                                    ICellStyle cellStyle = workbook.CreateCellStyle();
                                    // Set the fill foreground color (you can change this to any color you desire)
                                    cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index; // for example, light orange

                                    // Set the fill pattern
                                    cellStyle.FillPattern = FillPattern.SolidForeground;

                                    // Apply the style to the cell
                                    newCell.CellStyle = cellStyle;
                                }
                                else if (clientValue.Item3.Equals("red"))
                                {
                                    ICellStyle cellStyle = workbook.CreateCellStyle();
                                    // Set the fill foreground color (you can change this to any color you desire)
                                    cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index; // for example, light orange

                                    // Set the fill pattern
                                    cellStyle.FillPattern = FillPattern.SolidForeground;

                                    // Apply the style to the cell
                                    newCell.CellStyle = cellStyle;
                                }
                                else if (clientValue.Item3.Equals("darkyellow"))
                                {
                                    ICellStyle cellStyle = workbook.CreateCellStyle();
                                    // Set the fill foreground color (you can change this to any color you desire)
                                    cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.DarkYellow.Index; // for example, light orange

                                    // Set the fill pattern
                                    cellStyle.FillPattern = FillPattern.SolidForeground;

                                    // Apply the style to the cell
                                    newCell.CellStyle = cellStyle;
                                }

                                if (clientValue.Item1 != 0)
                                {
                                    newCell.SetCellValue("Розничная сумма: " + clientValue.Item1.ToString());
                                }

                                newCell = currentRow.GetCell(maxCellsCount + 3) ?? currentRow.CreateCell(maxCellsCount + 3); // Get the first cell or create a new one if it doesn't exist


                                if (clientValue.Item3.Equals("green"))
                                {
                                    ICellStyle cellStyle = workbook.CreateCellStyle();
                                    // Set the fill foreground color (you can change this to any color you desire)
                                    cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Green.Index; // for example, light orange

                                    // Set the fill pattern
                                    cellStyle.FillPattern = FillPattern.SolidForeground;

                                    // Apply the style to the cell
                                    newCell.CellStyle = cellStyle;
                                }
                                else if (clientValue.Item3.Equals("yellow"))
                                {
                                    ICellStyle cellStyle = workbook.CreateCellStyle();
                                    // Set the fill foreground color (you can change this to any color you desire)
                                    cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index; // for example, light orange

                                    // Set the fill pattern
                                    cellStyle.FillPattern = FillPattern.SolidForeground;

                                    // Apply the style to the cell
                                    newCell.CellStyle = cellStyle;
                                }
                                else if (clientValue.Item3.Equals("red"))
                                {
                                    ICellStyle cellStyle = workbook.CreateCellStyle();
                                    // Set the fill foreground color (you can change this to any color you desire)
                                    cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index; // for example, light orange

                                    // Set the fill pattern
                                    cellStyle.FillPattern = FillPattern.SolidForeground;

                                    // Apply the style to the cell
                                    newCell.CellStyle = cellStyle;
                                }
                                else if (clientValue.Item3.Equals("darkyellow"))
                                {
                                    ICellStyle cellStyle = workbook.CreateCellStyle();
                                    // Set the fill foreground color (you can change this to any color you desire)
                                    cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.DarkYellow.Index; // for example, light orange

                                    // Set the fill pattern
                                    cellStyle.FillPattern = FillPattern.SolidForeground;

                                    // Apply the style to the cell
                                    newCell.CellStyle = cellStyle;
                                }

                                newCell.SetCellValue(clientValue.Item2);

                                newCell = null;

                                diffPdfExclSum.Remove(fioKey);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }

        var pdfRow = sheet.LastRowNum + 15;
        int cnt = 1;
        foreach (var dict in diffPdfExclSum)
        {
            // Set the value of cell A1 to "New Value"
            IRow row = sheet.GetRow(pdfRow) ?? sheet.CreateRow(pdfRow); // Get the first row or create a new one if it doesn't exist
            NPOI.SS.UserModel.ICell cell = row.GetCell(0) ?? row.CreateCell(0);
            cell.SetCellValue(cnt);
            cell = row.GetCell(1) ?? row.CreateCell(1);// Get the first cell or create a new one if it doesn't exist
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Orange.Index; // for example,  orange
            cellStyle.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = cellStyle;
            cell.SetCellValue(dict.Key);

            cell = row.GetCell(2) ?? row.CreateCell(2); // Get the first cell or create a new one if it doesn't exist
            cellStyle = workbook.CreateCellStyle();
            cellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Orange.Index; // for example,  orange
            cellStyle.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = cellStyle;
            cell.SetCellValue(dict.Value.Item2);
            cnt++;
            pdfRow++;
        }

        IRow roww = sheet.GetRow(pdfRow++) ?? sheet.CreateRow(pdfRow++); // Get the first row or create a new one if it doesn't exist
        NPOI.SS.UserModel.ICell cellww = roww.GetCell(0) ?? roww.CreateCell(0);
        cellww.SetCellValue(firstReplase);
        cellww = roww.GetCell(1) ?? roww.CreateCell(1);// Get the first cell or create a new one if it doesn't exist
        cellww.SetCellValue(secondReplase);

        // Save the changes
        // using FileStream fileStream = new(excelFile.FileName, FileMode.Create, FileAccess.Write);
        // workbook.Write(fileStream);
        return workbook;
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
    }
    public List<PdfTable> ParseTableFromPage(string pageText)
    {
        List<PdfTable> rows = new();

        // Define regular expression patterns to match each row of the table
        string pattern = @"(\d+)\s+(\p{L}+\s+\p{L}+\s+\p{L}+)\s+(\d{12})\s+(\S+)\s+([\d,]+\.\d{2})";
        Regex regex = new(pattern, RegexOptions.Multiline);
        Regex periodPattern = new(@"Период\s*:\s*([\d.]+)\s*-\s*([\d.]+)");


        // Match rows using the regular expression pattern
        MatchCollection matches = regex.Matches(pageText);
        System.Text.RegularExpressions.Match periodMatch = periodPattern.Match(pageText);

        //if (periodMatch.Captures.Count != 0)
        //{
        //    StartDate = DateTime.Parse(periodMatch.Groups[1].Value);
        //    EndDate = DateTime.Parse(periodMatch.Groups[2].Value);
        //}

        // Extract data from each match and create PdfTable objects
        foreach (System.Text.RegularExpressions.Match match in matches.Cast<System.Text.RegularExpressions.Match>())
        {
            PdfTable row = new();
            row.Number = match.Groups[1].Value;
            row.Fio = match.Groups[2].Value;
            row.IIN = match.Groups[3].Value;
            row.AccountNumber = match.Groups[4].Value;
            desimalString = match.Groups[5].Value;
            firstReplase = match.Groups[5].Value.Replace(",", "");
            secondReplase = firstReplase.Replace(".", ",");
            decimalDecimal = Convert.ToDecimal(secondReplase);
            row.Amount = decimalDecimal;
            if (!String.IsNullOrEmpty(row.Fio)) row.Fio.ToUpper();
            rows.Add(row);
        }

        return rows;
    }
    public List<PdfTable> ExtractTextFromPdf(IFormFile pdfFile)
    {
        List<PdfTable> pdfdatas = new();
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
    public List<Models.Excel.ExcelRow> ExtractDataFromExcel(IFormFile excelFile)
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
        List<Models.Excel.ExcelRow> excelRows = new();

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
                    Models.Excel.ExcelRow excelRow = new();
                    string queue = currentRow.GetCell(queueRow)?.ToString() ?? "";
                    int convertedQueue = -1;

                    try
                    {
                        convertedQueue = Convert.ToInt32(queue);
                        if (convertedQueue != -1)
                        {
                            Console.WriteLine(convertedQueue);
                            excelRow.Number = queue;
                            excelRow.Fio = currentRow.GetCell(fioRow)?.ToString() ?? "";
                            string sum = currentRow.GetCell(transferredToTheBankCell)?.ToString() ?? "0";
                            excelRow.Amount = Convert.ToDecimal(sum);
                            excelRows.Add(excelRow);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }
        return excelRows;
    }
}
