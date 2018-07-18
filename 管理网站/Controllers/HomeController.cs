using Calculation.Dal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace 管理网站.Controllers
{
    public class HomeController : BaseController
    {
        /// <summary>
        /// 首页
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            if (this.CurrentUser!=null&&this.CurrentUser.IsAuthenticated)
            {
                return View();
            }
            else
            {
                return new RedirectToRouteResult(new RouteValueDictionary(new { controller = "account", action = "login", returnMessage = "您无权查看." }));
            }
        }
    }

     
}