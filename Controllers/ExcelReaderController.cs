
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace MigratorAzureDevops.Controllers
{
    public class ExcelReaderController : Controller
    {

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
                //Stream zipStream;
               // System.IO.Compression.ZipArchive zipArchive;
                //List<WorkitemFromExcel> WiList;
                ExcelPackage excel = new ExcelPackage(excelStream);
                //WIOps.ConnectWithPAT(URI, UserPAT);
                DT = ReadExcel(excel);
                
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
            return View();

        }
        public static DataTable ReadExcel(ExcelPackage Excel)
        {
            DataTable Dt = new DataTable();
            //Console.Write("Enter The Ecel File Path:");
            /*string ExcelPath=Console.ReadLine();*/
            foreach ( var sheets in Excel.Workbook.Worksheets)
            {
                var WorkSheet = sheets;

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
            }
            

            
            return Dt;
        }
    }
}