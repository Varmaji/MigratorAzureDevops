using log4net;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.OAuth;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MigratorAzureDevops.Class
{
    class WIOps
    {
        static string Url;
        public static string status = "";
        //public static int status = 0;
        static WorkItemTrackingHttpClient WitClient;
        public static ILog logger = LogManager.GetLogger("ErrorLog");
        public static WorkItem CreateWorkItem(string ProjectName, string WorkItemTypeName, Dictionary<string, object> Fields)
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();
            try
            {
                Dictionary<string, object> Fields1 = FormatDates(Fields);
                ////foreach(var key in Fields.Keys)
                ////{
                ////    if(key.Contains("Date"))
                ////    {
                ////        string Date=Fields[key].ToString();
                ////        DateTime oDate = Convert.ToDateTime(Date);
                ////        string str = oDate.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'");

                ////    }
                ////}
                foreach (var item in Fields1.Keys.ToList())

                    patchDocument.Add(new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/" + item,
                        Value = Fields1[item]
                    });

                
            }
            catch (Exception ex)
            {
                logger.Info(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + "\t" + ex.Message + "\n" + ex.StackTrace + "\n");
                throw (ex);
            }
            return WitClient.CreateWorkItemAsync(patchDocument, ProjectName, WorkItemTypeName).Result;
        }

        public static Dictionary<string, object> FormatDates(Dictionary<string, object> Fld)
        {
            Dictionary<string, object> DATFld = new Dictionary<string, object>();
            foreach(var i in Fld)
            {
                DATFld.Add(i.Key,i.Value);
            }

            foreach (var flds in DATFld.Keys)
            {
                if (flds.Contains("Date"))
                {
                    string Date = Fld[flds].ToString();
                    DateTime oDate = Convert.ToDateTime(Date, System.Globalization.CultureInfo.GetCultureInfo("hi-IN").DateTimeFormat);
                    string FormattedDate = oDate.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'");
                    Fld[flds] = FormattedDate;
                }
            }
            return Fld;
        }


        public static WorkItem UpdateWorkItemLink(int parentId, int childId, string message)
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();
            try
            {
                patchDocument.Add(new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/relations/-",
                    Value = new
                    {
                        rel = "System.LinkTypes.Hierarchy-Reverse",
                        url = Url + "/_apis/wit/workitems/" + parentId,
                        attributes = new
                        {
                            comment = "Linking the workitems"
                        }
                    }
                });

                return WitClient.UpdateWorkItemAsync(patchDocument, childId).Result;
            }
            catch (Exception E)
            {
                throw (E);
            }
        }
        public static WorkItem UpdateWorkItemFields(int WIId, Dictionary<string, object> Fields)
        {
            try
            {
                JsonPatchDocument patchDocument = new JsonPatchDocument();
                foreach (var key in Fields.Keys)
                {
                    JsonPatchOperation Jsonpatch = new JsonPatchOperation()
                    {
                        Operation = Operation.Add,
                        Path = "/fields/" + key,
                        Value = Fields[key]
                    };
                    if (!patchDocument.Contains(Jsonpatch))
                        patchDocument.Add(Jsonpatch);
                }
                if (patchDocument.Count != 0)
                    return WitClient.UpdateWorkItemAsync(patchDocument, WIId).Result;
                else
                    return null;
            }
            catch (Exception E)
            {
                throw (E);
            }
        }

        public static void ConnectWithPAT(string ServiceURL, string PAT)
        {
            try
            {
                Url = ServiceURL;
                VssConnection connection = new VssConnection(new Uri(ServiceURL), new VssOAuthAccessTokenCredential(PAT));
                InitClients(connection);
            }
            catch (Exception E)
            {
                throw E;
            }
        }
        static void InitClients(VssConnection Connection)
        {
            try
            {
                WitClient = Connection.GetClient<WorkItemTrackingHttpClient>();
            }
            catch (Exception E)
            {
                throw (E);
            }
        }
    }
}
