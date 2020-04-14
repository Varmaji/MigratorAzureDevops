﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigratorAzureDevops.Models
{
    public class WorkitemFromExcel
    {
        public int id { get; set; }
        public string tittle { get; set; }
        public ParentWorkItem parent { get; set; }
        public string WiState { get; set; }
        public string AreaPath { get; set; }
        public string Itertation { get; set; }
    }
    public class ParentWorkItem
    {
        public int Id { get; set; }
        public string tittle { get; set; }
    }
    public class WItypeStates
    {
        public List<States> value { get; set; }
    }
    public class States
    {
        public string name { get; set; }
    }
    public class WIS
    {
        public List<WI> WorkItems { get; set; }
    }
    public class WI
    {
        public string Id { get; set; }
    }
}
