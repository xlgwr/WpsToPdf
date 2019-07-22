using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace WpsToPdf.WebApi.Controllers
{
    public class HomeController : BaseController
    {
        public JsonResult Index()
        {
            ViewBag.Title = "Home Page";

            return ToJsonResult("你可以正常调用了。");
        }
    }
}
