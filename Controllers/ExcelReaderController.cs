using Microsoft.TeamFoundation.Common;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using MigratorAzureDevops.Class;
using MigratorAzureDevops.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace MigratorAzureDevops.Controllers
{
    public class ExcelReaderController : Controller
    {
        static Dictionary<string, DataTable> sheets;
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        static string BaseUrl = "";
        static string UserPAT;
        static string ProjectName;
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();
        static public string OrganisationName="";
        static public string OldTeamProject = "";

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

        static APIRequest req;

        [HttpPost]
        public ActionResult ReadExcelFile(HttpPostedFileBase Excel, string Organisation, string PAT, string SourceProj, string DestionationProj)
        {
            try
            {
                var excelStream = Excel.InputStream;
                ExcelPackage excel = new ExcelPackage(excelStream);

                ReadExcel(excel);

            }
            catch (IndexOutOfRangeException)
            {
                ViewBag.message = "No Work Sheets Found";
            }
            catch (Exception ex)
            {
                ViewBag.message = "Something Went Wrong, Please Download Excel/Attachments From 'Export Attachments'";
                throw (ex);
            }
            OrganisationName = Organisation;
            BaseUrl = "https://dev.azure.com/" + Organisation + "/";
            ProjectName = DestionationProj;
            UserPAT = PAT;
            WIOps.ConnectWithPAT(BaseUrl, UserPAT);
            req = new APIRequest(UserPAT);

            string response = req.ApiRequest(BaseUrl+ DestionationProj + "/_apis/wit/fields?api-version=5.1");
            Fields fieldsList = JsonConvert.DeserializeObject<Fields>(response);          
            List<SelectListItem> list = new List<SelectListItem>();
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
            sheets = new Dictionary<string, DataTable>();
            foreach (var WorkSheet in Excel.Workbook.Worksheets)
            {
                DataTable Dt = new DataTable();
                int rowCount = WorkSheet.Dimension.End.Row;
                int colCount = WorkSheet.Dimension.End.Column;
                DataRow row;
                for (int i = WorkSheet.Dimension.Start.Row; i <= rowCount; i++)
                {
                    row = Dt.NewRow();
                    for (int j = WorkSheet.Dimension.Start.Column; j <= colCount; j++)
                    {
                        string ColName = "";
                        if (i == WorkSheet.Dimension.Start.Row)
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
                            ColName = WorkSheet.Cells[WorkSheet.Dimension.Start.Row, j].Value.ToString();
                            if (WorkSheet.Cells[i, j].Value != null)
                                row[ColName] = WorkSheet.Cells[i, j].Value.ToString();
                        }
                    }
                    if (i != WorkSheet.Dimension.Start.Row)
                        Dt.Rows.Add(row);
                }
                int x = 1;
                if (sheets.ContainsKey(WorkSheet.Name))
                    sheets.Add(WorkSheet.Name + "(" + x++ + ")", Dt);
                else
                    sheets.Add(WorkSheet.Name, Dt);
            }        

        }

        static Dictionary<string, string> MappedFields;
        static List<string> SheetNames = new List<string>();

        [HttpPost]
        public JsonResult createExcel(Dictionary<string, string> FList, string SheetName)
        {
            try
            {
                DT = sheets[SheetName];
                foreach (var item in FList)
                {
                    if(DT.Columns.Contains(item.Key))
                    DT.Columns[item.Key].ColumnName = item.Value;
                }                
                if(SheetNames.Contains(SheetName))
                    return Json("WorkItems From This Sheet Already Migrated", JsonRequestBehavior.AllowGet);
                List<WorkitemFromExcel> WiList = GetWorkItems();
                CreateLinks(WiList);
                bool isUpdated = UpdateWIFields();
                if (!SheetNames.Contains(SheetName))
                    SheetNames.Add(SheetName);
                WIOps.status = "Successfully Migrated"+WiList.Count+" workitems";
                return Json(WIOps.status, JsonRequestBehavior.AllowGet);
            }
            catch (Exception E)
            {
                return Json(E.InnerException.ToString(), JsonRequestBehavior.AllowGet);
            }

        }
        public List<WorkitemFromExcel> GetWorkItems()
        {
            try
            {
                OldTeamProject = null;
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
                            if (OldTeamProject.IsNullOrEmpty())
                            {
                                if (!dr["Iteration Path"].ToString().IsNullOrEmpty())
                                {
                                    string ColVal = dr["Iteration Path"].ToString();
                                    string[] ValArr = ColVal.Split('/');
                                    OldTeamProject = ValArr[0];
                                }

                            }
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

                throw (E);

            }

        }
        public void CreateLinks(List<WorkitemFromExcel> WiList)
        {
            foreach (var wi in WiList)
            {
                WorkItem Wi;
                if (wi.parent != null)
                    Wi = WIOps.UpdateWorkItemLink(wi.parent.Id, wi.id, "");
            }

        }
        public ParentWorkItem getParentData(DataTable dt, int rowindex, int columnindex)
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
                throw E;
             
            }

        }

        static int createWorkItem(DataRow Dr)
        {
            try
            {                 
                Dictionary<string, object> fields = new Dictionary<string, object>();
                foreach (DataColumn column in DT.Columns)
                {
                    if (!string.IsNullOrEmpty(Dr[column].ToString()))
                    {
                        if (column.ToString().StartsWith("Title")||column.ToString()=="CLM_ID")
                            fields.Add("Title", Dr[column].ToString());
                    }
                }
                Object Wiql = new { query = "Select  [Id] From WorkItems Where [System.Title] = '" + fields["Title"] + "' AND  [System.Team Project] ='"+ProjectName+ "' AND  [Custom.CLM_ID] ='" + fields["CLM_ID"] + "'" };
                string response = req.ApiRequest(BaseUrl+"/_apis/wit/wiql?api-version=4.1", "POST", Wiql.ToString());
                WIS ExistingWI = JsonConvert.DeserializeObject<WIS>(response);
                if(ExistingWI.WorkItems.Count>=0)
                    return int.Parse(ExistingWI.WorkItems[0].Id);
                var newWi = WIOps.CreateWorkItem(ProjectName, Dr["Work Item Type"].ToString(), fields);
                return newWi.Id.Value;
            }
            catch (Exception E)
            {
                throw E;
            }
        }
        public static bool UpdateWIFields()
        {
            try
            {
                foreach (DataRow row in DT.Rows)
                {
                    Dictionary<string, object> Updatefields = new Dictionary<string, object>();
                    foreach (DataColumn col in DT.Columns)
                    {
                        if (!string.IsNullOrEmpty(row[col].ToString()))
                        {
                            if (col.ToString() != "ID" && col.ToString() != "Reason")
                            {
                                string val = row
                                    [col.ToString()].ToString().Replace(OldTeamProject, ProjectName).TrimStart('\\');
                                if (!string.IsNullOrEmpty(val))
                                {                                   
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
                throw E;
            }

        }
    }
}