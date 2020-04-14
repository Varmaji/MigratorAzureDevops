using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MigratorAzureDevops.Controllers
{
    public class WelcomeController : Controller
    {
        
        public ActionResult Welcomepage()
        {
            return View();
        }
    }
}