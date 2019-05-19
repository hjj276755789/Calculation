using Calculation.Base;
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
        private FW_QXGL_DataProvider _fw = new FW_QXGL_DataProvider();
        /// <summary>
        /// 首页
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            if (this.CurrentUser!=null&&this.CurrentUser.IsAuthenticated)
            {
                this.ViewBag.jsxx = _fw.GET_JSLB(this.CurrentUser.YHBH);
                return View();
            }
            else
            {
                return new RedirectToRouteResult(new RouteValueDictionary(new { controller = "account", action = "login", returnMessage = "您无权查看." }));
            }
        }
        /// <summary>
        /// 获取数据任务详情
        /// </summary>
        /// <returns></returns>
        public JsonResult GET_Z_DATA_TASK_INFO()
        {
            //传递参数：人员编号，年份，周次
            var cjba = _fw.GET_Z_DATA_TASK_INFO_CJBA(DateTime.Now.Year, Base_date.GET_Z_of_Y(DateTime.Now));
            var xzys = _fw.GET_Z_DATA_TASK_INFO_XZYS(DateTime.Now.Year, Base_date.GET_Z_of_Y(DateTime.Now));
            var tdcj = _fw.GET_Z_DATA_TASK_INFO_TDCJ(DateTime.Now.Year, Base_date.GET_Z_of_Y(DateTime.Now));
            var rgsj = _fw.GET_Z_DATA_TASK_INFO_RGSJ(DateTime.Now.Year, Base_date.GET_Z_of_Y(DateTime.Now));
            var obj = new
            {
                cjbh = cjba == 0 ? 0 : 1,
                xzys = xzys == 0 ? 0 : 1,
                tdcj = tdcj == 0 ? 0 : 1,
                rgsj = rgsj == 0 ? 0 : 1
            };
            return Json(obj);
        }
        /// <summary>
        /// 获取周报任务详情
        /// </summary>
        /// <returns></returns>
        public JsonResult GET_ZB_TASK_INFO()
        {
            var obj = _fw.GET_ZB_TASK_INFO(CurrentUser.YHBH, DateTime.Now.Year, Base_date.GET_Z_of_Y(DateTime.Now));
            return Json(obj);
        }
    }

     
}