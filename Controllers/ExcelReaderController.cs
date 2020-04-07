
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
using MigratorAzureDevops.Class;
using Newtonsoft.Json;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;

namespace MigratorAzureDevops.Controllers
{
    public class ExcelReaderController : Controller
    {
        static Dictionary<string, DataTable> sheets;
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        static string BaseUrl= "https://dev.azure.com/";
        static string UserPAT;// = "qbi2it66pkjvlj7p4whh7efbkdjqzemume5xazf7ogspqmcieosa";
        static string ProjectName;// = "Agile Project";//"HOLMES-TrainingStudio";
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();       
        static public string OldTeamProject;// = "HOLMES-AutomationStudio";
        
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
        public ActionResult ReadExcelFile(HttpPostedFileBase Excel, string Organisation, string PAT,string SourceProj,string DestionationProj)
        {
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
            BaseUrl += Organisation;
            ProjectName = DestionationProj;
            OldTeamProject = SourceProj;
            UserPAT = PAT;
            WIOps.ConnectWithPAT(BaseUrl, UserPAT);
            APIRequest req = new APIRequest(UserPAT);
            string response=req.ApiRequest("https://dev.azure.com/"+Organisation+"/"+DestionationProj+"/_apis/wit/fields?api-version=5.1");
            Fields fieldsList = JsonConvert.DeserializeObject<Fields>(response);
            /*var model = new sheetList()
            {
                Sheets = sheets,
                fields=fieldsList.value
            };
            string data = JsonConvert.SerializeObject(model);
            ViewBag.model = data;*/
            List<SelectListItem> list =new List<SelectListItem>();
            foreach (var key in sheets.Keys)
            {
                list.Add(new SelectListItem() { Text = key, Value = JsonConvert.SerializeObject(sheets[key]) });
            }
            List<SelectListItem> flist = new List<SelectListItem>();
            foreach (var field in fieldsList.value)
            {
                flist.Add(new SelectListItem() { Text = field.name, Value = field.name });
            }
            ViewBag.fields = flist;
            ViewBag.Selectlist = list;
                return View("SheetsDrop");

        }

        
        public void ReadExcel(ExcelPackage Excel)
        {
            //Console.Write("Enter The Ecel File Path:");
            /*string ExcelPath=Console.ReadLine();*/
            sheets  = new Dictionary<string, DataTable>();
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

        static Dictionary<string, string> MappedFields ;

        [HttpPost]
        public JsonResult createExcel(Dictionary<string, string> FList,string SheetName)
        {
            string status = "";
            try
            {


                MappedFields = FList;
                DT = sheets[SheetName];
                List<WorkitemFromExcel> WiList = GetWorkItems();
                if (WiList.Count <= 0)
                    return Json("Some Error Has Occured", JsonRequestBehavior.AllowGet);
                CreateLinks(WiList);
                bool isUpdated=UpdateWIFields();
                if (isUpdated==true)
                    status = "Successfully Migrated" + DT.Rows.Count + " WorkItems";
                else
                    status = "Something Went Wrong";
                return Json(status, JsonRequestBehavior.AllowGet);
            }catch(Exception E)
            {
                return Json(E.InnerException, JsonRequestBehavior.AllowGet);
            }

        }
        public  List<WorkitemFromExcel> GetWorkItems()
        {
            try
            {
                List<WorkitemFromExcel> workitemlist = new List<WorkitemFromExcel>();
                if (DT.Rows.Count > 0)
                {
                    for (int i = 0; i < DT.Rows.Count; i++)
                    {
                        DataRow dr = DT.Rows[i];
                        WorkitemFromExcel item = new WorkitemFromExcel();
                        if (DT.Rows[i] != null)
                        {
                            item.id = createWorkItem(dr);
                            //("WorkItemPublish Created= " + item.id);
                            dr["ID"] = item.id.ToString();


                            int columnindex = 0;
                            foreach (var col in TitleColumns)
                            {
                                if (!string.IsNullOrEmpty(col))
                                {
                                    if (!string.IsNullOrEmpty(dr[col].ToString()))
                                    {
                                        item.tittle = dr[col].ToString();
                                        if (i > 0 && columnindex > 0)
                                            item.parent = getParentData(DT, i - 1, columnindex);
                                        break;
                                    }
                                }
                                columnindex++;
                            }
                            workitemlist.Add(item);
                        }

                    }
                }
                return workitemlist;
            }
            catch (Exception E)
            {
                throw(E);
                return null;
            }

        }
        public  void CreateLinks(List<WorkitemFromExcel> WiList)
        {
            foreach (var wi in WiList)
            {
                WorkItem Wi;
                if (wi.parent != null)
                    Wi=WIOps.UpdateWorkItemLink(wi.parent.Id, wi.id, "");
                
            }
           
        }
        public  ParentWorkItem getParentData(DataTable dt, int rowindex, int columnindex)
        {
            try
            {
                ParentWorkItem workItem = new ParentWorkItem();

                if (columnindex > 0)
                {
                    for (int i = rowindex; i >= 0; i--)
                    {

                        DataRow dr = dt.Rows[i];
                        int colindex = columnindex;
                        while (colindex > 0)
                        {
                            int index = colindex - 1;
                            if (!string.IsNullOrEmpty(dr[TitleColumns[index]].ToString()))
                            {
                                workItem.Id = int.Parse(dr["ID"].ToString());
                                workItem.tittle = dr[TitleColumns[index]].ToString();
                                break;
                            }
                            colindex--;
                        }
                        if (!string.IsNullOrEmpty(workItem.tittle))
                        { break; }

                    }
                }
                return workItem;
            }
            catch (Exception E)
            {
                throw(E);
                return null;
            }

        }


        static int createWorkItem(DataRow Dr)
        {
            Dictionary<string, object> fields = new Dictionary<string, object>();
            foreach (DataColumn column in DT.Columns)
            {
                if (!string.IsNullOrEmpty(Dr[column].ToString()))
                {
                    if (column.ToString().StartsWith("Title"))
                        fields.Add("Title", Dr[column].ToString());
                }
            }
            var newWi = WIOps.CreateWorkItem(ProjectName, Dr["Work Item Type"].ToString(), fields);            
            return newWi.Id.Value;
        }
        public static bool UpdateWIFields()
        {
            try
            {
                foreach (DataRow row in DT.Rows)
                {
                    //Throw("Updating Fields of" + row["ID"]);
                    Dictionary<string, object> Updatefields = new Dictionary<string, object>();
                    foreach (DataColumn col in DT.Columns)
                    {
                        if (!string.IsNullOrEmpty(row[col].ToString()))
                        {
                            if (col.ToString() != "ID" && col.ToString() != "Reason" && col.ToString() != "Work Item Type" && !col.ToString().StartsWith("Title"))
                            {
                                string val = row
                                    [col.ToString()].ToString().Replace(OldTeamProject, ProjectName).TrimStart('\\');
                                if (!string.IsNullOrEmpty(val))
                                {
                                    if (MappedFields.ContainsKey(col.ToString()))
                                        Updatefields.Add(MappedFields[col.ToString()], val);
                                    else
                                        Updatefields.Add(col.ToString(), val);
                                }
                            }
                        }
                    }
                    WIOps.UpdateWorkItemFields(int.Parse(row["ID"].ToString()), Updatefields);
                }
                return true;
            }
            catch (Exception E)
            {
                return false;
            }

        }
    }
}