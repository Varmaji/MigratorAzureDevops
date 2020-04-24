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
using Microsoft.TeamFoundation.Common;
using System.Configuration;
using distribution_copy.Models;
using MigratorAzureDevops.Service;
using System.Threading;
using System.Web.UI;
using log4net;

namespace MigratorAzureDevops.Controllers
{

    public class ExcelReaderController : Controller
    {

        static Dictionary<string, DataTable> sheets;
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        static string BaseUrl = "https://dev.azure.com/";
        static string UserPAT;// = "qbi2it66pkjvlj7p4whh7efbkdjqzemume5xazf7ogspqmcieosa";
        static string ProjectName;// = "Agile Project";//"HOLMES-TrainingStudio";
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();
        static public string OldTeamProject;// = "HOLMES-AutomationStudio";
        static public string OrganizationName;
        public static int workitemCount;
        public static ILog logger = LogManager.GetLogger("ErrorLog");
        public static int WorkitemCount
        {
            get
            {

                return workitemCount;
            }
            set
            {
                workitemCount = value;
            }
        }
        public int WICount;


        //This delegate is implemented to for fetching workitem count created everytime
        private delegate string[] ProcessEnvironment(string SheetName);



        // GET: ExcelReader
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Validation()
        {
            try
            {
                Session["visited"] = "1";
                string url = "https://app.vssps.visualstudio.com/oauth2/authorize?client_id={0}&response_type=Assertion&state=User1&scope={1}&redirect_uri={2}";
                string redirectUrl = System.Configuration.ConfigurationManager.AppSettings["RedirectUri"];
                string clientId = System.Configuration.ConfigurationManager.AppSettings["ClientId"];
                string AppScope = System.Configuration.ConfigurationManager.AppSettings["appScope"];
                url = string.Format(url, clientId, AppScope, redirectUrl);
                return Redirect(url);
            }
            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.Message + "\n" + ex.StackTrace + "\n");
                return RedirectToAction("Welcomepage", "Welcome");

            }

        }

        [HttpGet]
        public ActionResult ReadExcelFile()
        {
            Thread.Sleep(2000);
            AccountService Account = new AccountService();
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
                    List<SelectListItem> OrganizationList = new List<SelectListItem>();
                    foreach (var i in accountList.value)
                    {
                        OrganizationList.Add(new SelectListItem { Text = i.accountName, Value = i.accountName });
                    }
                    ViewBag.OrganizationList = OrganizationList;
                }
                catch (Exception ex)
                {
                    logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                    logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                    throw (ex);
                }
            }

            return View();
        }

        [HttpPost]
        public ActionResult ReadExcelFile(HttpPostedFileBase Excel/*, string Organisation, string PAT, string SourceProj, string DestionationProj*/)
        {
            //if (Session["PAT"] == null)
            //{
            //    RedirectToAction("");
            //}
            //PAT = Session["PAT"].ToString();
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
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                ViewBag.message = "Something Went Wrong, Please Download Excel/Attachments From 'Export Attachments'";

            }
            //BaseUrl += Organisation;
            //OrganizationName = Organisation;
            //ProjectName = DestionationProj;
            ////OldTeamProject = SourceProj;
            //UserPAT = PAT;
            //WIOps.ConnectWithPAT(BaseUrl + Organisation + "/", UserPAT);
            return RedirectToAction("Destin", "ExcelReader", new { ErrorMessage = ViewBag.message });
        }

        [HttpPost]
        public JsonResult ProjectList(string ORG)
        {
            AccountService service = new AccountService();
            var pm = service.GetApi<ProjectModel>("https://dev.azure.com/" + ORG + "/_apis/projects?api-version=5.1");
            return Json(pm.Value, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult Destin()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Destin(string Organisation, string DestionationProj)
        {
            try
            {
                if (Session["PAT"] == null)
                {
                    RedirectToAction("");
                }
                string PAT = Session["PAT"].ToString();
                OrganizationName = Organisation;
                ProjectName = DestionationProj;
                //OldTeamProject = SourceProj;
                UserPAT = PAT;
                WIOps.ConnectWithPAT(BaseUrl + Organisation + "/", UserPAT);
            }
            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                ViewBag.ErrorMessage = ex.Message;
            }
            return RedirectToAction("SheetsDrop", "ExcelReader");
        }

        public JsonResult AccountList()
        {
            AccountsResponse.AccountList accountList = new AccountsResponse.AccountList();
            if (Session["AccountList"] != null)
            {
                accountList = (AccountsResponse.AccountList)Session["AccountList"];
            }
            return Json(accountList.value, JsonRequestBehavior.AllowGet);
        }
        public void ReadExcel(ExcelPackage Excel)
        {
            //Console.Write("Enter The Ecel File Path:");
            /*string ExcelPath=Console.ReadLine();*/
            sheets = new Dictionary<string, DataTable>();
            try
            {
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

            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                throw (ex);
            }
            /*return sheets;*/

        }

        public ActionResult SheetsDrop()
        {
            try
            {
                APIRequest req = new APIRequest(UserPAT);
                string response = req.ApiRequest("https://dev.azure.com/" + OrganizationName + "/" + ProjectName + "/_apis/wit/fields?api-version=5.1");
                Fields fieldsList = JsonConvert.DeserializeObject<Fields>(response);

                //Read Excel sheets Column value 
                List<SelectListItem> list = new List<SelectListItem>();
                foreach (var key in sheets.Keys)
                {
                    list.Add(new SelectListItem() { Text = key, Value = JsonConvert.SerializeObject(sheets[key]) });
                }

                //Read Azure Devops field name
                List<SelectListItem> flist = new List<SelectListItem>();
                foreach (var field in fieldsList.value)
                {
                    flist.Add(new SelectListItem() { Text = field.name, Value = field.name });
                }

                ViewBag.fields = flist;
                ViewBag.Selectlist = list;
            }
            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                ViewBag.ErrorMessage = ex.Message;
            }
            return View();
        }


        //static Dictionary<string, string> MappedFields;
        static List<string> SheetNames = new List<string>();

        [HttpPost]
        public bool createExcel(Dictionary<string, string> FList, string SheetName)
        {
            WorkitemCount = 0;
            try
            {
                //if(SheetNames.Contains(SheetName))
                //    return Json("WorkItems From This Sheet Already Migrated", JsonRequestBehavior.AllowGet);
                //MappedFields = FList;
                DT = sheets[SheetName];

                foreach (var item in FList)
                {
                    if (DT.Columns.Contains(item.Key))
                        DT.Columns[item.Key].ColumnName = item.Value;
                }
                ProcessEnvironment task = new ProcessEnvironment(GetWorkItems);
                task.BeginInvoke(SheetName, new AsyncCallback(EndMethod), task);

                return true;
            }
            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                ViewBag.ErrorMessage = ex.Message;
                return false;
            }
            //finally
            //{
            //    ExportDataSetToExcel();
            //}

        }

        public void EndMethod(IAsyncResult result)
        {
            try
            {
                ProcessEnvironment processTask = (ProcessEnvironment)result.AsyncState;
                string[] strResult = processTask.EndInvoke(result);
            }
            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                throw (ex);
            }
        }

        static DataSet DS = new DataSet();
        //public void ExportDataSetToExcel()
        //{
        //    DataTable people = (DataTable)Session["people"];

        //    // Open excel file and insert data table.
        //    ExcelPackage ef = new ExcelPackage();
        //    ef.LoadXls(Server.MapPath("MyData.xls"));
        //    ExcelWorksheet ws = ef.Worksheets[0];
        //    ws.InsertDataTable(people, "A1", true);

        //    Response.Clear();

        //    // Stream file to browser, in required type.
        //    switch (this.RadioButtonList1.SelectedValue)
        //    {
        //        case "XLS":
        //            Response.ContentType = "application/vnd.ms-excel";
        //            Response.AddHeader("Content-Disposition", "attachment; filename=" +
        //                 "Report.xls");
        //            ef.SaveXls(Response.OutputStream);
        //            break;

        //        case "XLSX":
        //            Response.ContentType = "application/vnd.openxmlformats";
        //            Response.AddHeader("Content-Disposition", "attachment; filename=" +
        //                 "Report.xlsx");
        //            // With XLSX it is a bit more complicated as MS Packaging API
        //            // can't write directly to Response.OutputStream.
        //            // Therefore we use temporary MemoryStream.
        //            MemoryStream ms = new MemoryStream();
        //            ef.SaveXlsx(ms)
        //            ms.WriteTo(Response.OutputStream);
        //            break;
        //    }
        //    Response.End();
        //}

        public ContentResult WorkItemCount(int WI)
        {
            return Content(WorkitemCount.ToString());
        }

        APIRequest req = new APIRequest(UserPAT);

        public string[] GetWorkItems(string SheetName)
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
                            WorkitemCount += 1;


                            if (OldTeamProject.IsNullOrEmpty())
                            {

                                if (!dr["Area Path"].ToString().IsNullOrEmpty() || !dr["Iteration Path"].ToString().IsNullOrEmpty())
                                {

                                    string ColVal = dr["Iteration Path"].ToString();
                                    string[] ValArr = ColVal.Split('/');
                                    OldTeamProject = ValArr[0];

                                }

                            }
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
                CreateLinks(workitemlist);
                bool isUpdated = UpdateWIFields();
                WIOps.status = "Successfully Migrated workitems";
                //WIOps.status = WorkitemCount;
                if (!SheetNames.Contains(SheetName))
                    SheetNames.Add(SheetName);
                return new string[] { "success" };
            }

            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                throw (ex);

            }
            //return new string[] { "" };

        }
        public void CreateLinks(List<WorkitemFromExcel> WiList)
        {
            try
            {
                foreach (var wi in WiList)
                {

                    WorkItem Wi;
                    if (wi.parent != null)
                        Wi = WIOps.UpdateWorkItemLink(wi.parent.Id, wi.id, "");

                }
            }
            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
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
            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                return null;

            }

        }
        public int createWorkItem(DataRow Dr)
        {
            try
            {
                Dictionary<string, object> fields = new Dictionary<string, object>();
                foreach (DataColumn column in DT.Columns)
                {
                    if (column.ColumnName.StartsWith("Iteration") || column.ColumnName.StartsWith("Area"))
                    {
                        if (OldTeamProject.IsNullOrEmpty())
                        {
                            string ColVal = Dr[column.ColumnName].ToString();
                            string[] ValArr;
                            if (ColVal.Contains("/"))
                                ValArr = ColVal.Split('/');
                            else
                                ValArr = ColVal.Split('\\');
                            OldTeamProject = ValArr[0];
                        }

                    }
                    if (!string.IsNullOrEmpty(Dr[column].ToString()))
                    {
                        if (column.ToString().StartsWith("Title"))
                            fields.Add("Title", Dr[column].ToString());
                        else if (column.ToString() != "State" && column.ToString() != "Reason" && column.ToString() != "ID")
                        {
                            if (column.ColumnName.StartsWith("Iteration") || column.ColumnName.StartsWith("Area"))
                            {
                                if (!string.IsNullOrEmpty(OldTeamProject))
                                {
                                    string val = Dr[column.ToString()].ToString().Replace(OldTeamProject, ProjectName).TrimStart('\\');
                                    fields.Add(column.ToString(), val);
                                }
                            }
                            else
                                fields.Add(column.ToString(), Dr[column].ToString());
                        }

                    }
                }
                Object Wiql = new { query = "Select  [Id] From WorkItems Where [System.Title] = '" + fields["Title"] + "' AND  [System.TeamProject] ='" + ProjectName + "'" };
                string response = req.ApiRequest(BaseUrl +  OrganizationName + "/_apis/wit/wiql?api-version=4.1", "POST", JsonConvert.SerializeObject(Wiql));
                WIS ExistingWI = JsonConvert.DeserializeObject<WIS>(response);
                if (ExistingWI.WorkItems.Count > 0)
                {
                    return int.Parse(ExistingWI.WorkItems[0].Id);
                }
                else
                {
                    var newWi = WIOps.CreateWorkItem(ProjectName, Dr["Work Item Type"].ToString(), fields);
                    logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                    logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "WorkItem ID Created: " + newWi.Id.Value);
                    return newWi.Id.Value;
                }

            }

            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                throw (ex);
            }

            //return 0;
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
                                    //if (MappedFields.ContainsKey(col.ToString()))
                                    //    Updatefields.Add(MappedFields[col.ToString()], val);
                                    //else
                                    Updatefields.Add(col.ToString(), val);
                                }
                            }
                        }
                    }
                    var Fieldsss = WIOps.FormatDates(Updatefields);
                    WIOps.UpdateWorkItemFields(int.Parse(row["ID"].ToString()), Fieldsss);

                }

                return true;
            }
            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + "Project Name: " + ProjectName + "\t Organization Selected: " + OrganizationName);
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.InnerException.Message);
                throw (ex);
            }

        }

        public FileResult Download()
        {
            var directory = new DirectoryInfo(Server.MapPath("~\\Logs"));
            var myFile = (from f in directory.GetFiles()
                          orderby f.LastWriteTime descending
                          select f).First();
            string FilePath = Path.Combine(directory.ToString(), myFile.ToString());
            //byte[] fileBytes = System.IO.File.ReadAllBytes(FilePath);
            string fileName = "Logs.txt";
            return File(FilePath, "text/plain", fileName);
        }
    }
}