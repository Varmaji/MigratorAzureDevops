using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;
using MigratorAzureDevops;
using Newtonsoft.Json;
using MigratorAzureDevops.Models;

namespace MigratorAzureDevops
{
    public class Operations
    {
        static string Url;//= "https://dev.azure.com/Organisation";
        static string UserPAT;
        static string ProjectName;
        static public int titlecount = 0;
        static public List<string> titles = new List<string>();
        static DataTable DT;
        static List<string> TitleColumns = new List<string>();
        static public string OldTeamProject = "HOLMES-TrainingStudio";
        static string ExcelPath;
        static APIRequest apiClient;
        public static DataTable ReadExcel(Excel._Worksheet xlWorksheet)
        {
            DataTable Dt = new DataTable();

            try
            {
               /* Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"" + ExcelPath);
                foreach (Excel.Worksheet wSheet in xlWorkbook.Worksheets)
                { }
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];*/
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                DataRow row;
               // Dictionary<string, List<States>> WITypestates = new Dictionary<string, List<States>>();

                string ColName = "";
                for (int i = 1; i <= rowCount; i++)
                {
                    row = Dt.NewRow();
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (i == 1)
                        {
                            ColName = xlRange.Cells[j][i].Value.ToString();
                            if (ColName.StartsWith("Title"))
                            {
                                TitleColumns.Add(ColName);
                            }
                            DataColumn column = new DataColumn(ColName);
                            Dt.Columns.Add(column);

                            continue;
                        }
                        ColName = xlRange.Cells[j][1].Value.ToString();
                        if (xlRange.Cells[j][i].Value != null)
                        {
                            string val = xlRange.Cells[j][i].Value.ToString().TrimStart('\\');
                            //val = val.Replace(OldTeamProject, ProjectName);
                            /*if (ColName == "WorkItem Type")
                            {
                                if (!WItypeStates.ContainsKey(val))
                                {
                                    string url = Url + "/" + ProjectName + "/_apis/wit/workitemtypes/" + val + "/states?api-version=5.1-preview.1";
                                    string response = apiClient.ApiRequest(url);
                                    WItypeStates states = JsonConvert.DeserializeObject<WItypeStates>(response);
                                    WITypestates.Add(val, states.value);
                                }
                            }*/
                            row[ColName] = val;
                        }
                    }
                    if (i != 1)
                    {
                       /* string Witype = row["WorkItem Type"].ToString();
                        States state = new States() { name = row["State"].ToString() };*/
                       /* if (!WITypestates[Witype].Contains(state))*/
                            Dt.Rows.Add(row);
                    }
                    /*string teststring =row.ItemArray[3].ToString();*/
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return Dt;
        }
    }
}