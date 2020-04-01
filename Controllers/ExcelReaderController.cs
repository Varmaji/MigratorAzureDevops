
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using MigratorAzureDevops.Models;
namespace MigratorAzureDevops.Controllers
{
    public class ExcelReaderController : Controller
    {
        static Dictionary<string, DataTable> sheets = new Dictionary<string, DataTable>();
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        // GET: ExcelReader
        public ActionResult Index()
        {
            return View();
        }
        
        [HttpGet]
        public ActionResult ReadExcelFile()
        {
            return View();
        }

        [HttpPost]
        public ActionResult ReadExcelFile(HttpPostedFileBase Excel)
        {
            //URI = @"https://dev.azure.com/" + Org + "/";
            //UserPAT = Session["PAT"] != null ? Session["PAT"].ToString() : "";
            //ProjectName = Proj;
            try
            {
                var excelStream = Excel.InputStream;              
                ExcelPackage excel = new ExcelPackage(excelStream);
                //WIOps.ConnectWithPAT(URI, UserPAT);
                ReadExcel(excel);
                
            }
            catch (IndexOutOfRangeException)
            {
                ViewBag.message = "No Work Sheets Found";
            }
            catch (Exception ex)
            {
                //throw ex;
                ViewBag.message = "Something Went Wrong, Please Download Excel/Attachments From 'Export Attachments'";

            }
            var model = new sheetList()
            {
                Sheets = sheets
            }; 
            return Json(model,JsonRequestBehavior.AllowGet);

        }
        public void ReadExcel(ExcelPackage Excel)
        {
            //Console.Write("Enter The Ecel File Path:");
            /*string ExcelPath=Console.ReadLine();*/
            foreach ( var WorkSheet in Excel.Workbook.Worksheets)
            {
                DataTable Dt = new DataTable();
                int rowCount = WorkSheet.Dimension.End.Row;
                int colCount = WorkSheet.Dimension.End.Column;                
                DataRow row;
                for (int i = 1; i <= rowCount; i++)
                {
                    row = Dt.NewRow();
                    for (int j = 1; j <= colCount; j++)
                    {
                        string ColName;
                        if (i == 1)
                        {
                            ColName = WorkSheet.Cells[i, j].Value.ToString();
                            if (ColName.StartsWith("Title"))
                            {
                                TitleColumns.Add(ColName);
                            }
                            DataColumn column = new DataColumn(ColName);
                            Dt.Columns.Add(column);
                        }
                        else
                        {
                            ColName = WorkSheet.Cells[1, j].Value.ToString();
                            if (WorkSheet.Cells[i, j].Value != null)
                                row[ColName] = WorkSheet.Cells[i, j].Value.ToString();
                        }
                    }
                    if (i != 1)
                        Dt.Rows.Add(row);
                }
                int x = 1;
                if(sheets.ContainsKey(WorkSheet.Name))
                sheets.Add(WorkSheet.Name+"("+x++ +")", Dt);
                else
                    sheets.Add(WorkSheet.Name, Dt);
            }
            /*return sheets;*/
        }
    }
}