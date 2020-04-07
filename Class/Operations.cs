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
    }
}