using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace MigratorAzureDevops.Models
{
    public class ExcelModel
    {
        public List<string> sheetName { get; set; }

        public string path { get; set; }
    }
    public class sheetList
    {
        public Dictionary<string,DataTable> Sheets { get; set; }
        public string  selectedSheet { get; set; }
        public List<Field> fields { get; set; }
    }
    public class Fields
    {
        public int count { get; set; }
        public List<Field> value { get; set; }

    }
    public class Field
    {
        public string name { get; set; }
    }
    public class inputmodel
    {
        public Dictionary<string,string> MappedFields { get;set; }
        public string Organisation { get; set; }
        public string SourceProj { get; set; }
        public string DestinationProj { get; set; }
        public string PAT { get; set; }
        public string SheetName { get; set; }

    }
}