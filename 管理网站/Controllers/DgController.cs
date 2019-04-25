using Calculation.Dal;
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
        Dg_DataProvider _dd = new Dg_DataProvider();
        #region 页面

       
        // GET: Dg
        public ActionResult Index()
        {
            return View();
        }
        #endregion

        #region 数据

        public JsonResult DG_Grid(string tj,string nf,string zc,int pagesize,int pagenow)
        {
            var obj = _dd.GET_DG(tj, nf, zc, pagesize, pagenow);
            var s = new
            {
                pagenow = obj.PageNumber,
                datacount = obj.TotalPageCount,
                d = obj
            };
            return Json(s);
        }
        #endregion
    }
}