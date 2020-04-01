using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MigratorAzureDevops.Models
{
    public class ExcelModel
    {
        public List<string> sheetName { get; set; }

        public string path { get; set; }
    }
}