
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
using MigratorAzureDevops.Models.Accounts;
using System.Configuration;
using MigratorAzureDevops.Services;

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
        readonly AccountService Account = new AccountService();

        // GET: ExcelReader
        public ActionResult Index()
        {
            
            return View();
        }
        
        public ActionResult ReadExcelFile()
        {
            if (Session["visited"] == null)
                return RedirectToAction("../Account/Verify");

            if (Session["PAT"] == null)
            {
                try
                {
                    AccessDetails _accessDetails = new AccessDetails();
                    AccountsResponse.AccountList accountList = null;
                    string code = Session["PAT"] == null ? Request.QueryString["code"] : Session["PAT"].ToString();
                    string redirectUrl = ConfigurationManager.AppSettings["RedirectUri"];
                    string clientId = ConfigurationManager.AppSettings["ClientSecret"];
                    string accessRequestBody = string.Empty;
                    accessRequestBody = Account.GenerateRequestPostData(clientId, code, redirectUrl);
                    _accessDetails = Account.GetAccessToken(accessRequestBody);
                    ProfileDetails profile = Account.GetProfile(_accessDetails);
                    if (!string.IsNullOrEmpty(_accessDetails.access_token))
                    {
                        Session["PAT"] = _accessDetails.access_token;

                        if (profile.displayName != null || profile.emailAddress != null)
                        {
                            Session["User"] = profile.displayName ?? string.Empty;
                            Session["Email"] = profile.emailAddress ?? profile.displayName.ToLower();
                        }
                    }
                    accountList = Account.GetAccounts(profile.id, _accessDetails);
                    Session["AccountList"] = accountList;
                    string pat = Session["PAT"].ToString();
                    
                }
                catch (Exception) { }
            }
            return View();
        }

        [HttpPost]
        public ActionResult ReadExcelFile(HttpPostedFileBase Excel)
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

            return RedirectToAction("SheetsDrop","ExcelReader");
        }
        static string Organisation = "";
        [HttpPost]
        public ActionResult SheetsDrop(string OrganisationName, string SourceProj, string DestionationProj)
        {
            ProjectName = DestionationProj;
            OldTeamProject = SourceProj;
            Organisation = OrganisationName;
            UserPAT = Session["PAT"].ToString();
            WIOps.ConnectWithPAT(BaseUrl + Organisation, UserPAT);
            APIRequest req = new APIRequest(UserPAT);
            string response = req.ApiRequest(BaseUrl + Organisation + "/" + ProjectName + "/_apis/wit/fields?api-version=5.1");
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
            return View();
        }

        [HttpGet]
        public ActionResult SheetsDrop()
        {
            List<SelectListItem> list = new List<SelectListItem>();
            foreach (var key in sheets.Keys)
            {
                list.Add(new SelectListItem() { Text = key, Value = JsonConvert.SerializeObject(sheets[key]) });
            }
            ViewBag.Selectlist = list;

            return View();
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
                for (int i = WorkSheet.Dimension.Start.Row; i <= rowCount; i++)
                {
                    row = Dt.NewRow();
                    for (int j = WorkSheet.Dimension.Start.Column; j <= colCount; j++)
                    {
                        string ColName="";
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
           
            try
            {
                MappedFields = FList;
                DT = sheets[SheetName];
                List<WorkitemFromExcel> WiList = GetWorkItems();               
                CreateLinks(WiList);
                bool isUpdated=UpdateWIFields();
                WIOps.status="Successfully Migrated workitems";
                return Json(WIOps.status, JsonRequestBehavior.AllowGet);
            }catch(Exception E)
            {
                return Json(E.InnerException.ToString(), JsonRequestBehavior.AllowGet);
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
                return null;
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
                        if (column.ToString().StartsWith("Title"))
                            fields.Add("Title", Dr[column].ToString());
                    }
                }
                var newWi = WIOps.CreateWorkItem(ProjectName, Dr["Work Item Type"].ToString(), fields);
                return newWi.Id.Value;
            }
            catch(Exception E)
            {
                return -1;
            }
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
                throw (E);
            }

        }
    }
}