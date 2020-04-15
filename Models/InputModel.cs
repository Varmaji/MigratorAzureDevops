using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace distribution_copy.Models
{
    public class InputModel
    {
        public string OrganizationName { get; set; }
        public string ProjectName { get; set; }
        public string WorkItemType { get; set; }
        public string AssignedTo { get; set; }
        public string Sprint { get; set; }
        public string CreatedDate { get; set; }
        public string State { get; set; }
    }
}