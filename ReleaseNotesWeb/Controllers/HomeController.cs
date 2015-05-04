using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ReleaseNotesWeb.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "TFS Release Notes";
            return View();
        }

        public ActionResult Waiting()
        {
            ViewBag.Title = "Waiting";
            return PartialView("Waiting");
        }
    }
}
