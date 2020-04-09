using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MigratorAzureDevops.Models
{
    public class ProjectModel
    {
        public int Count { get; set; }
        public List<ProjectDetails> Value { get; set; }
    }
    public class ProjectDetails
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }
}