using Microsoft.Office.Interop.Excel;

using MigratorAzureDevops.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace MigratorAzureDevops.Controllers
{
    public class OpenExcelController : Controller
    {
        List<string> WorksheetName = new List<string>();
        ExcelModel excel = new ExcelModel();
        // GET: OpenExcel
        public ActionResult Index()
        {
            return View();
            
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile)
        {
            ExcelFields fields = new ExcelFields();
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select a excel file<br>";
                return View("Index");
            }
            else
            {
                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {

                    string path = Server.MapPath("~/ExcelFile/" + excelfile.FileName);
                   
                   
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);

                    


                    //Read data from Excel file
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook =application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    excel.sheetName=LoadExcelWorkSheets(path);



                    return View(excel);
                }
                else
                {
                    ViewBag.Error = "File Type is incorrect<br>";
                    return View("Index");
                }

            }

        }

        private List<string> LoadExcelWorkSheets(string path)
        {
            string filePath =path ;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook excelBook = xlApp.Workbooks.Open(filePath);
            String[] excelSheets = new String[excelBook.Worksheets.Count];
            int i = 0;
            
            foreach (Excel.Worksheet wSheet in excelBook.Worksheets)
            {
               WorksheetName.Add(wSheet.Name);
                i++;
            }
            return WorksheetName;
        }
    }
}