﻿using Calculation.Dal;
using Calculation.Models.Enums;
using Calculation.Models;
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
            this.ViewBag.jslb = _fw.GET_JSLB();
            this.ViewBag.yhjslb = _fw.GET_JSLB(yhid);
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

        public PartialViewResult jsqxgl(int jsid)
        {
            this.ViewBag.jsid = jsid;
            this.ViewBag.qxlb = _fw.GET_GQXLB();
            this.ViewBag.jsqxlb = _fw.GET_QXLB(jsid);
            return PartialView();
        }
        #endregion


        #region 数据块
        #region 用户块
        public JsonResult GET_YHLB()
        {
            return Json(_fw.GET_YHLB());
        }
        #endregion

        

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


        #region 角色块
        public JsonResult GET_JSLB()
        {
            return Json(_fw.GET_JSLB());
        }

        public JsonResult Remove_YHJS(int yhid, int jsid)
        {
            if (_fw.DEL_YHJS(yhid, jsid))
                return Json(SResult.Success);
            else return Json(SResult.Error("设置失败"));
        }

        public JsonResult ADD_YHJS(int yhid, int jsid)
        {
            if (_fw.ADD_YHJS(yhid, jsid))
                return Json(SResult.Success);
            else return Json(SResult.Error("设置失败"));
        }
        #endregion

        #region 权限块
        public JsonResult ADD_JSQX(int jsid, int gqxid)
        {
            if(_fw.ADD_JSQX(jsid, gqxid))
                return Json(SResult.Success);
            else return Json(SResult.Error("设置失败"));
        }
        public JsonResult Remove_JSQX(int jsid,int gqxid)
        {
            if (_fw.DEL_JSQX(jsid, gqxid))
                return Json(SResult.Success);
            else return Json(SResult.Error("设置失败"));
        }
        
        #endregion
        #endregion

    }
}