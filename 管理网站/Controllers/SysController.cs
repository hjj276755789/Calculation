using Calculation.Dal;
using Calculation.Models.Enums;
using Calculation.Models.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace 管理网站.Controllers
{
  
    public class SysController : BaseController
    {
        private FW_QXGL_DataProvider _fw;
        public SysController()
        {
            _fw = new FW_QXGL_DataProvider();
        }

        #region 页面块

        [IdentityCheck]
        // GET: Sys
        public ActionResult Index()
        {
            return View();
        }
        [IdentityCheck]
        public ActionResult yhgl()
        {
            return View();
        }
        [IdentityCheck]
        [HttpGet]
        public PartialViewResult add_yhxx()
        {
            return PartialView();
        }

        [IdentityCheck]
        public PartialViewResult yhjs(int yhid)
        {
            this.ViewBag.yhid = yhid;
            return PartialView();
        }

        [IdentityCheck]
        public ActionResult jsgl()
        {
            return View();
        }
        [IdentityCheck]
        public ActionResult qxgl()
        {
            return View();
        }
        #endregion


        #region 数据块
        public JsonResult GET_YHLB()
        {
            return Json(_fw.GET_YHLB());
        }
        public JsonResult GET_JSLB()
        {
            return Json(_fw.GET_JSLB());

        }
        public JsonResult GET_JSLB_BY_ID(int yhid)
        {
            return Json(_fw.GET_JSLB(yhid));
        }
        public JsonResult GET_QXLB()
        {
            return Json(_fw.GET_QXLB());
        }
        public JsonResult GET_QXLB_BY_ID(int jsid)
        {
            return Json(_fw.GET_QXLB(jsid));
        }
        [HttpPost]
        public JsonResult ADD_YHXX(string yhm,string yhmm,string cfmm)
        {
            if (string.IsNullOrEmpty(yhm)||string.IsNullOrEmpty(yhmm) ||yhmm != cfmm)
                return Json(SResult.Error("用户名或密码不符合规范"));
            else
            {
                if (_fw.ADD_USER(yhm, yhmm, YH_LX.普通账号))
                    return Json(SResult.Success);
                else
                    return Json(SResult.Error("新增用户失败！"));

            }
        }
        [HttpPost]
        public JsonResult DEL_YHXX(string id)
        {

                if (_fw.DEL_USER(id))
                    return Json(SResult.Success);
                else
                    return Json(SResult.Error("新增用户失败！"));
        }
        #endregion

    }
}