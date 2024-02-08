using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Linq;
using ConverterAndChecker.Models;

namespace ConverterAndChecker.Controllers
{
    public class FileUploadController : Controller
    {
        [HttpGet]
        public ActionResult UploadFiles()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadFiles(IFormFile pdfFile, IFormFile excelFile)
        {
            UploadedFilesViewModel model = new();
            model.PdfFiles = new List<string>();
            model.ExcelFiles = new List<string>();

            if (pdfFile != null && pdfFile.Length > 0)
            {
                string pdfFileName = Path.GetFileName(pdfFile.FileName);
                model.PdfFiles.Add(pdfFileName);
            }

            if (excelFile != null && excelFile.Length > 0)
            {
                string excelFileName = Path.GetFileName(excelFile.FileName);
                model.ExcelFiles.Add(excelFileName);
            }

            return View(model);
        }


        public ActionResult ComparisonResult()
        {
            // Logic to prepare comparison results
            return View();
        }
    }

}
