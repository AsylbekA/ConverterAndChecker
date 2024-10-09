namespace ConverterAndChecker.Models;

public class UploadViewModel
{
    public IFormFile PdfFile { get; set; }
    public IFormFile XlsxFile { get; set; }
    public string TextPattern { get; set; }
}
