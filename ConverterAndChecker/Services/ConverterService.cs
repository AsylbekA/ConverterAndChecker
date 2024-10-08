using ConverterAndChecker.Models.Pdf;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Globalization;
using System.Text.RegularExpressions;
using NPOI.HPSF;
using ConverterAndChecker.Models.Excel;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace ConverterAndChecker.Services
{
    public class ConverterService
    {
        public string fileName;
        private DateTime StartDate;
        private DateTime EndDate;

        string desimalString = "";
        decimal decimalDecimal;
        string firstReplase;
        string secondReplase;
        public List<PdfTable> ExtractTextFromPdf(IFormFile pdfFile,int cnt)
        {
            List<PdfTable> pdfdatas = new();
            using (PdfReader reader = new(pdfFile.OpenReadStream()))
            {
                for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber++)
                {
                    string pageText = PdfTextExtractor.GetTextFromPage(reader, pageNumber);
                    if (cnt == 0)
                    {
                        pdfdatas.AddRange(ParseTableInshuranceFromPage(pageText));
                    }
                    else
                    {
                        pdfdatas.AddRange(ParseTableInshuranceFIOFromPage(pageText));
                    }
                }
            }
            return pdfdatas;
        }

        public List<PdfTable> ParseTableInshuranceFromPage(string pageText)
        {
            List<PdfTable> rows = new();

            // Define regular expression patterns to match each row of the table
            string pattern = @"(\d+)\s+(\p{L}+\s+\p{L}+\s+\p{L}+)\s+(\d{12})\s+([\d,]+\.\d{2})";
            //string pattern = @"(\d+)\s+(\p{L}+\s+\p{L}+\s+\p{L}+)\s+(\d{12})\s+(\S+)\s+([\d,]+\.\d{2})";
            //string pattern = @"(\d+)\s+([\p{Cyrillic}]+\s+[\p{Cyrillic}]+\s+[\p{Cyrillic}]+)\s+(\d{12})\s+(\S+)\s+([\d,]+\.\d{2})";

            Regex regex = new(pattern, RegexOptions.Multiline);
            Regex periodPattern = new(@"Период\s*:\s*([\d.]+)\s*-\s*([\d.]+)");


            // Match rows using the regular expression pattern
            MatchCollection matches = regex.Matches(pageText);
            //System.Text.RegularExpressions.Match periodMatch = periodPattern.Match(pageText);

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
                //row.AccountNumber = match.Groups[4].Value;
                desimalString = match.Groups[4].Value;

                //firstReplase = match.Groups[6].Value;
                // Define culture info with appropriate settings
                CultureInfo culture = new CultureInfo("en-US");

                // Specify NumberStyles to handle commas and periods
                NumberStyles style = NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands;

                // Convert string to decimal
                decimal result = decimal.Parse(desimalString, NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
                secondReplase = result.ToString();
                row.Amount = result;
                if (!String.IsNullOrEmpty(row.Fio)) row.Fio.ToUpper();
                rows.Add(row);
            }

            return rows;
        }

        public List<PdfTable> ParseTableInshuranceFIOFromPage(string pageText)
        {
            List<PdfTable> rows = new();

            // Define regular expression patterns to match each row of the table
            string pattern = @"(\d+)\s+(\p{L}+\s+\p{L}+)\s+(\d{12})\s+([\d,]+\.\d{2})";
            //string pattern = @"(\d+)\s+(\p{L}+\s+\p{L}+\s+\p{L}+)\s+(\d{12})\s+(\S+)\s+([\d,]+\.\d{2})";
            //string pattern = @"(\d+)\s+([\p{Cyrillic}]+\s+[\p{Cyrillic}]+\s+[\p{Cyrillic}]+)\s+(\d{12})\s+(\S+)\s+([\d,]+\.\d{2})";

            Regex regex = new(pattern, RegexOptions.Multiline);
            Regex periodPattern = new(@"Период\s*:\s*([\d.]+)\s*-\s*([\d.]+)");


            // Match rows using the regular expression pattern
            MatchCollection matches = regex.Matches(pageText);
            //System.Text.RegularExpressions.Match periodMatch = periodPattern.Match(pageText);

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
                //row.AccountNumber = match.Groups[4].Value;
                desimalString = match.Groups[4].Value;

                //firstReplase = match.Groups[6].Value;
                // Define culture info with appropriate settings
                CultureInfo culture = new CultureInfo("en-US");

                // Specify NumberStyles to handle commas and periods
                NumberStyles style = NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands;

                // Convert string to decimal
                decimal result = decimal.Parse(desimalString, NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
                secondReplase = result.ToString();
                row.Amount = result;
                //if (!String.IsNullOrEmpty(row.Fio)) row.Fio;
                rows.Add(row);
            }

            return rows;
        }

        public List<Models.Excel.ExcelRow> ExtractInshuranceFromExcel(IFormFile excelFile)
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
            int iinRow = -1;
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
                                if (val.Equals("Взносы на обязательное медицинское страхование"))
                                {
                                    transferredToTheBankCell = col;
                                }
                                else if (val.Equals(queueIcon))
                                {
                                    queueRow = col;
                                }
                                else if (val.Equals("Фамилия имя отчество"))
                                {
                                    fioRow = col;
                                }else if (val.Equals("ИИН"))
                                {
                                    iinRow = col;
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
                                string iin = currentRow.GetCell(iinRow)?.ToString() ?? "";
                                if (iin.Length == 12)
                                {
                                    Console.WriteLine(convertedQueue);
                                    excelRow.Number = queue;
                                    excelRow.Fio = currentRow.GetCell(fioRow)?.ToString() ?? ""; ;
                                    excelRow.Iin = iin;
                                    string sum = currentRow.GetCell(transferredToTheBankCell)?.ToString() ?? "0";
                                    excelRow.Amount = Convert.ToDecimal(sum);
                                    excelRows.Add(excelRow);
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
            return excelRows;
        }
        public List<Models.Excel.ExcelRow> ExtractENPFFromExcel(IFormFile excelFile)
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
            int iinRow = -1;
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
                                if (val.Equals("ОПВ"))
                                {
                                    transferredToTheBankCell = col;
                                }
                                else if (val.Equals(queueIcon))
                                {
                                    queueRow = col;
                                }
                                else if (val.Equals("Фамилия имя отчество"))
                                {
                                    fioRow = col;
                                }
                                else if (val.Equals("ИИН"))
                                {
                                    iinRow = col;
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
                                string iin = currentRow.GetCell(iinRow)?.ToString() ?? "";
                                if (iin.Length == 12)
                                {
                                    Console.WriteLine(convertedQueue);
                                    excelRow.Number = queue;
                                    excelRow.Fio = currentRow.GetCell(fioRow)?.ToString() ?? ""; ;
                                    excelRow.Iin = iin;
                                    string sum = currentRow.GetCell(transferredToTheBankCell)?.ToString() ?? "0";
                                    excelRow.Amount = Convert.ToDecimal(sum);
                                    excelRows.Add(excelRow);
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
            return excelRows;
        }
        public List<PdfTable> ParseTableFromPage515(string pageText)
        {
            List<PdfTable> rows = new();

            // Define regular expression patterns to match each row of the table

            string pattern = @"(\d+)\s+(\p{L}+\s+\p{L}+\s+\p{L}+)\s+(\d{12})\s+(\S+)\s+([\d,]+\.\d{2})";

            Regex regex = new(pattern, RegexOptions.Multiline);
            Regex periodPattern = new(@"Период\s*:\s*([\d.]+)\s*-\s*([\d.]+)");


            // Match rows using the regular expression pattern
            MatchCollection matches = regex.Matches(pageText);
            //System.Text.RegularExpressions.Match periodMatch = periodPattern.Match(pageText);

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

                firstReplase = match.Groups[6].Value;
                // Define culture info with appropriate settings
                CultureInfo culture = new CultureInfo("en-US");

                // Specify NumberStyles to handle commas and periods
                NumberStyles style = NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands;

                // Convert string to decimal
                decimal result = decimal.Parse(desimalString, NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands, CultureInfo.InvariantCulture);
                secondReplase = result.ToString();
                row.Amount = result;
                if (!String.IsNullOrEmpty(row.Fio)) row.Fio.ToUpper();
                rows.Add(row);
            }

            return rows;
        }
        public List<Models.Excel.ExcelRow> Extract515FromExcel(IFormFile excelFile)
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

        public MemoryStream SetExcelInshurance(Dictionary<string, ExcelRows> excelKeyValuePairs, Dictionary<string, PdfTables> pdfKeyValuePairs)
        {
            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                // Add a new worksheet to the Excel package
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add header
                worksheet.Cells[1, 1].Value = "№ п/п";
                worksheet.Cells[1, 2].Value = "Фамилия Имя Отчество";
                worksheet.Cells[1, 3].Value = "ИИН";
                worksheet.Cells[1, 4].Value = "Сумма в мед страх";
                worksheet.Cells[1, 5].Value = "Сумма в Ведомость";
                worksheet.Cells[1, 6].Value = "Розница";
                worksheet.Cells[1, 7].Value = "Причина";

                using (var headerRange = worksheet.Cells[1, 1, 1, 7])
                {
                    headerRange.Style.Font.Color.SetColor(Color.Black); // Черный цвет шрифта
                    headerRange.Style.Font.Bold = true; // Жирный шрифт
                    headerRange.Style.Font.Size = 12; // Увеличенный размер шрифта
                    headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Центрирование текста
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray); // Цвет фона заголовка
                }


                Dictionary<string, (decimal, string, string)> diffPdfExclSum = new Dictionary<string, (decimal, string, string)>();
                int row = 2;
                foreach (var val in excelKeyValuePairs)
                {
                    string upperkey = val.Key.ToUpper();
                    worksheet.Cells[row, 1].Value = val.Value.ExcelRow[0].Number;
                    if (val.Value.ExcelRow.Count > 1)
                    {
                        worksheet.Cells[row, 1].Value = val.Value.ExcelRow[0].Number + " и " + val.Value.ExcelRow[1].Number;
                    }
                    worksheet.Cells[row, 2].Value = val.Value.ExcelRow[0].Fio;
                    worksheet.Cells[row, 3].Value = val.Key;
                    worksheet.Cells[row, 5].Value = val.Value.Amount;
                    string color;
                    if (pdfKeyValuePairs.ContainsKey(upperkey))
                    {
                        var temp = pdfKeyValuePairs[upperkey];

                        worksheet.Cells[row, 4].Value = temp.Amount;
                        if (!diffPdfExclSum.ContainsKey(upperkey))
                        {
                            decimal diffSum = val.Value.Amount - temp.Amount;


                            string comment;
                            if (diffSum > 0)
                            {
                                comment = "Сумма из Расчетная ведомоста повышает сумму из Выписка по мед страх, по проведенным платежам.  Розница повышение = " + diffSum + "; Сумма из Расчетная ведомоста = " + val.Value.Amount + "; Сумма из мед страх, по проведенным платежам = " + temp.Amount;
                                color = "yellow";
                                worksheet.Cells[row, 6, row, 6].Style.Font.Color.SetColor(Color.Black);
                            }
                            else if (diffSum < 0)
                            {
                                comment = "Сумма из Мед страх, по проведенным платежам нехватает на сумму  Расчетная ведомоста.  Недостаточная сумма = " + diffSum + "; Сумма из Расчетная ведомоста  = " + val.Value.Amount + "; Сумма из мед страх, по проведенным платежам = " + temp.Amount;
                                color = "darkyellow";
                                worksheet.Cells[row, 6, row, 6].Style.Font.Color.SetColor(Color.DarkOrange);
                            }
                            else
                            {
                                comment = "";
                                color = "green";
                                worksheet.Cells[row, 6, row, 6].Style.Font.Color.SetColor(Color.Green);
                            }
                            worksheet.Cells[row, 6].Value = diffSum;
                            worksheet.Cells[row, 7].Value = comment;
                            diffPdfExclSum.Add(val.Key, (diffSum, comment, color));
                        }
                    }
                    else
                    {
                        worksheet.Cells[row, 7].Value = "Не найден клиент из списка мед страх, по проведенным платежам";
                        color = "red";
                        worksheet.Cells[row, 7, row, 7].Style.Font.Color.SetColor(Color.Red);
                        diffPdfExclSum.Add(val.Key, (val.Value.Amount, "Не найден клиент из Выписка, по проведенным платежам", "red"));
                    }


                    pdfKeyValuePairs.Remove(upperkey);

                    //worksheet.Cells[row, 1, row, 6].Style.Font.Color.;
                    row++;
                }

                row += 4;
                foreach (var val in pdfKeyValuePairs)
                {
                    worksheet.Cells[row, 1].Value = val.Value.PdfTable[0].Number;
                    if (val.Value.PdfTable.Count > 1)
                    {
                        worksheet.Cells[row, 1].Value = val.Value.PdfTable[0].Number + " и " + val.Value.PdfTable[1].Number;
                    }
                    worksheet.Cells[row, 2].Value = val.Value.PdfTable[0].Fio;
                    worksheet.Cells[row, 3].Value = val.Key;
                    worksheet.Cells[row, 4].Value = val.Value.Amount;
                    worksheet.Cells[row, 7].Value = "Не найден клиент из списка Расчетная ведомоста";
                    diffPdfExclSum.Add(val.Key, (val.Value.Amount, "Не найден клиент из Расчетная ведомоста", "white"));
                    row++;
                }
                package.SaveAs(stream);
                stream.Position = 0;

            }
            return stream;
        }

        public MemoryStream SetExcelOPV(Dictionary<string, ExcelRows> excelKeyValuePairs, Dictionary<string, PdfTables> pdfKeyValuePairs)
        {
            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                // Add a new worksheet to the Excel package
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add header
                worksheet.Cells[1, 1].Value = "№ п/п";
                worksheet.Cells[1, 2].Value = "Фамилия Имя Отчество";
                worksheet.Cells[1, 3].Value = "ИИН";
                worksheet.Cells[1, 4].Value = "Сумма РВ";
                worksheet.Cells[1, 5].Value = "Сумма 5-15А";
                worksheet.Cells[1, 6].Value = "Розница";
                worksheet.Cells[1, 7].Value = "Причина";

                using (var headerRange = worksheet.Cells[1, 1, 1, 7])
                {
                    headerRange.Style.Font.Color.SetColor(Color.Black); // Черный цвет шрифта
                    headerRange.Style.Font.Bold = true; // Жирный шрифт
                    headerRange.Style.Font.Size = 12; // Увеличенный размер шрифта
                    headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center; // Центрирование текста
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray); // Цвет фона заголовка
                }


                Dictionary<string, (decimal, string, string)> diffPdfExclSum = new Dictionary<string, (decimal, string, string)>();
                int row = 2;
                foreach (var val in excelKeyValuePairs)
                {
                    string upperkey = val.Key.ToUpper();
                    worksheet.Cells[row, 1].Value = val.Value.ExcelRow[0].Number;
                    if (val.Value.ExcelRow.Count > 1)
                    {
                        worksheet.Cells[row, 1].Value = val.Value.ExcelRow[0].Number + " и " + val.Value.ExcelRow[1].Number;
                    }
                    worksheet.Cells[row, 2].Value = val.Value.ExcelRow[0].Fio;
                    worksheet.Cells[row, 3].Value = val.Key;
                    worksheet.Cells[row, 4].Value = val.Value.Amount;
                    string color;
                    if (pdfKeyValuePairs.ContainsKey(upperkey))
                    {
                        var temp = pdfKeyValuePairs[upperkey];

                        worksheet.Cells[row, 5].Value = temp.Amount;
                        if (!diffPdfExclSum.ContainsKey(upperkey))
                        {
                            decimal diffSum = val.Value.Amount - temp.Amount;


                            string comment;
                            if (diffSum > 0)
                            {
                                comment = "Сумма из РВ повышает сумму из 5-15А, по проведенным платежам.  Розница повышение = " + diffSum;
                                color = "yellow";
                                worksheet.Cells[row, 6, row, 6].Style.Font.Color.SetColor(Color.Black);
                            }
                            else if (diffSum < 0)
                            {
                                comment = "Сумма из 5-15А недостаточно на сумму РВ.  Недостаточная сумма = " + diffSum;
                                color = "darkyellow";
                                worksheet.Cells[row, 6, row, 6].Style.Font.Color.SetColor(Color.DarkOrange);
                            }
                            else
                            {
                                comment = "";
                                color = "green";
                                worksheet.Cells[row, 6, row, 6].Style.Font.Color.SetColor(Color.Green);
                            }
                            worksheet.Cells[row, 6].Value = diffSum;
                            worksheet.Cells[row, 7].Value = comment;
                            diffPdfExclSum.Add(val.Key, (diffSum, comment, color));
                        }
                    }
                    else
                    {
                        worksheet.Cells[row, 7].Value = "Не найден клиент из списка 5-15А, по проведенным платежам";
                        color = "red";
                        worksheet.Cells[row, 7, row, 7].Style.Font.Color.SetColor(Color.Red);
                        diffPdfExclSum.Add(val.Key, (val.Value.Amount, "Не найден клиент из Выписка, по проведенным платежам", "red"));
                    }


                    pdfKeyValuePairs.Remove(upperkey);

                    //worksheet.Cells[row, 1, row, 6].Style.Font.Color.;
                    row++;
                }

                row += 4;
                foreach (var val in pdfKeyValuePairs)
                {
                    worksheet.Cells[row, 1].Value = val.Value.PdfTable[0].Number;
                    if (val.Value.PdfTable.Count > 1)
                    {
                        worksheet.Cells[row, 1].Value = val.Value.PdfTable[0].Number + " и " + val.Value.PdfTable[1].Number;
                    }
                    worksheet.Cells[row, 2].Value = val.Value.PdfTable[0].Fio;
                    worksheet.Cells[row, 3].Value = val.Key;
                    worksheet.Cells[row, 5].Value = val.Value.Amount;
                    worksheet.Cells[row, 7].Value = "Не найден клиент из списка РВ";
                    diffPdfExclSum.Add(val.Key, (val.Value.Amount, "Не найден клиент из Расчетная ведомоста", "white"));
                    row++;
                }
                package.SaveAs(stream);
                stream.Position = 0;

            }
            return stream;
        }
    }
}
