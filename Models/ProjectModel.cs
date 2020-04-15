using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MigratorAzureDevops.Models
{
    public class ProjectModel
    {
        public int Count { get; set; }
        //public string value { get; set; }
        public List<ProjectDetails> Value { get; set; }
        public List<ProjectDetails> WorkItems { get; set; }
        public ProjectDetails ProjectDetails { get; set; }
    }

    public class AuthoredBy
    {
        public string descriptor { get; set; }
    }

    public class ProjectDetails
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Url { get; set; }
        public string State { get; set; }
        public int Revision { get; set; }
        public string Visibility { get; set; }
        public DateTime LastUpdateTime { get; set; }
        //public orgCounts counts { get; set; }
        public AuthoredBy authoredBy { get; set; }
        public string originId { get; set; }


    }
}