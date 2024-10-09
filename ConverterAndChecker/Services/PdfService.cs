using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using System.Text;

namespace ConverterAndChecker.Services
{
    public class PdfService
    {
        public static string ExtractTextFromPdf(string path)
        {
            using (PdfReader reader = new PdfReader(path))
            using (PdfDocument pdfDoc = new PdfDocument(reader))
            {
                StringBuilder text = new StringBuilder();
                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    var page = pdfDoc.GetPage(i);
                    string ss = PdfTextExtractor.GetTextFromPage(page);
                    text.Append(ss);
                }
                return text.ToString();
            }
        }
    }
}
