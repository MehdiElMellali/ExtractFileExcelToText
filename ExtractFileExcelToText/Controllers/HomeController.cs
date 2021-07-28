using ExtractFileExcelToText.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExtractFileExcelToText.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

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

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ExtractFile(IFormFile file, CancellationToken cancellationToken)
        {
            StringBuilder sb = new StringBuilder();

            if (file == null || file.Length <= 0)
            {
            }

            if (!Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
            }


            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream, cancellationToken);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                    {
                        sb.AppendLine("");

                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells[row, j].Value != null)
                            {
                                sb.Append(worksheet.Cells[row, j].Value.ToString());
                                sb.Append(" ");
                            }
                        }
                    }
                }
            }

            var str = sb.ToString();
            ViewData["str"] = str;
            return View();
        } 
        
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> ExtractInformationByCelleAndRow(IFormFile file,string rowcellNum, int cellNum, int rowNum, CancellationToken cancellationToken)
        {
            StringBuilder sb = new StringBuilder();

            if (file == null || file.Length <= 0)
            {
            }

            if (!Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
            }


            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream, cancellationToken);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                        sb.AppendLine(worksheet.Cells[rowNum,cellNum].Text);
                        sb.AppendLine(worksheet.Cells[rowcellNum].Text);
                }
            }

            var str = sb.ToString();
            ViewData["str"] = str;
            return View();
        }
    }
}
