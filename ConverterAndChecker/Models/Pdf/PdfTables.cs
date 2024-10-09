namespace ConverterAndChecker.Models.Pdf;

public class PdfTables
{
    public decimal Amount { get; set; }
    public string FullInfo { get; set; }

    public List<PdfTable> PdfTable { get; set; }
}
