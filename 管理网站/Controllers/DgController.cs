using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace 管理网站.Controllers
{

    /// <summary>
    /// 定稿
    /// </summary>
    public class DgController : Controller
    {
        #region 页面

       
        // GET: Dg
        public ActionResult Index()
        {
            return View();
        }
        #endregion

        #region 数据

        public JsonResult DG_Grid(string tj,int pagesize,int pagenow)
        {
            return Json("");
        }
        #endregion
    }
}