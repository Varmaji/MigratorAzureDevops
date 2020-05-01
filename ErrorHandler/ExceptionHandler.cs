using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MigratorAzureDevops.ErrorHandler
{
    public class ExceptionHandler : FilterAttribute, IExceptionFilter
    {
        public void OnException(ExceptionContext filterContext)
        {
            //if (filterContext.Exception is KeyNotFoundException)
            //{

            //}

            //else if (filterContext.Exception is ColumnAlreadyExist)
            //{

            //}

            //else if (filterContext.Exception is WrongPATException)
            //{

            //}
            filterContext.Result = new ViewResult()
            {
                ViewName = "Error"
            };
            filterContext.ExceptionHandled = true;
        }

    }

    //public class ColumnAlreadyExist : Exception
    //{
    //    public ColumnAlreadyExist()
    //    {

    //    }

    //    public ActionResult WrongPATException(string message)
    //    {
    //        return RedirectResult("")
    //    }
    //}
}