using Calculation.Dal;
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
        private FW_KFS_DataProvider _kfs;
        public SysController()
        {
            _fw = new FW_QXGL_DataProvider();
            _kfs = new FW_KFS_DataProvider();
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
        public PartialViewResult update_yhmm(string yhbh)
        {
            this.ViewBag.yhbh = yhbh;
            return PartialView();
        }
        [IdentityCheck]
        public PartialViewResult yhjs(int yhbh)
        {
            this.ViewBag.yhbh = yhbh;
            this.ViewBag.jslb = _fw.GET_JSLB();
            this.ViewBag.yhjslb = _fw.GET_JSLB(yhbh);
            return PartialView();
        }

        [IdentityCheck]
        public ActionResult jsgl()
        {
            return View();
        }

        [IdentityCheck]
        public ActionResult fzkfs(string yhbh)
        {
            this.ViewBag.yhbh = yhbh;
            return View();
        }
        public PartialViewResult add_js()
        {
            return PartialView();
        }


        public PartialViewResult jsqxgl(int jsbh)
        {
            this.ViewBag.jsbh = jsbh;
            this.ViewBag.qxlb = _fw.GET_GQXLB();
            this.ViewBag.jsqxlb = _fw.GET_QXLB(jsbh);
            return PartialView();
        }

        public ActionResult kfs()
        {
            return View();
        }

        public PartialViewResult add_kfs()
        {
            return PartialView();
        }

        public ActionResult kfs_mb(string kfsbh)
        {
            this.ViewBag.kfsbh = kfsbh;
            return View();
        }
        #endregion


        #region 数据块
        #region 用户块
        public JsonResult GET_YHLB(string tj,int pagesize,int pagenow)
        {
            var obj = _fw.GET_YHLB(tj,pagesize, pagenow);
            var s = new
            {
                pagenow = obj.PageNumber,
                datacount = obj.TotalPageCount,
                d = obj
            };
            return Json(s);
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
        public JsonResult DEL_YHXX(string yhbh)
        {

                if (_fw.DEL_USER(yhbh))
                    return Json(SResult.Success);
                else
                    return Json(SResult.Error("新增用户失败！"));
        }
        [IdentityCheck]
        [HttpPost]
        public JsonResult Update_yhmm_xx(string yhbh,string yhmm)
        {

            if (_fw.update_yhmm_xx(yhbh, yhmm))
                return Json(SResult.Success);
            else
                return Json(SResult.Error("修改密码失败！"));
        }
        #endregion

        #region 角色块
        public JsonResult GET_JSLB()
        {
            return Json(_fw.GET_JSLB());
        }
        public JsonResult ADD_JSXX(string jsmc,string jsms)
        {
            if (_fw.ADD_JSXX(jsmc, jsms))
                return Json(SResult.Success);
            else return Json(SResult.Error("添加失败！"));
        }

        public JsonResult Remove_YHJS(int yhbh, int jsbh)
        {
            if (_fw.DEL_YHJS(yhbh, jsbh))
                return Json(SResult.Success);
            else return Json(SResult.Error("设置失败"));
        }

        public JsonResult ADD_YHJS(int yhbh, int jsbh)
        {
            if (_fw.ADD_YHJS(yhbh, jsbh))
                return Json(SResult.Success);
            else return Json(SResult.Error("设置失败"));
        }
        #endregion

        #region 权限块
        public JsonResult ADD_JSQX(int jsbh, int fqxbh)
        {
            if(_fw.ADD_JSQX(jsbh, fqxbh))
                return Json(SResult.Success);
            else return Json(SResult.Error("设置失败"));
        }
        public JsonResult Remove_JSQX(int jsbh,int fqxbh)
        {
            if (_fw.DEL_JSQX(jsbh, fqxbh))
                return Json(SResult.Success);
            else return Json(SResult.Error("设置失败"));
        }
        #endregion

        #region 开发商
        public JsonResult GET_KFSLB(string tj, int pagesize, int pagenow)
        {
            var obj = _kfs.FIND_KFS_XX(tj,pagesize,pagenow);
            var s = new
            {
                pagenow = obj.PageNumber,
                datacount = obj.TotalPageCount,
                d = obj
            };
            return Json(s);
        }


        [HttpPost]
        public JsonResult ADD_KFS(string kfsmc, string kfslxr, string kfslxrdh, string bz)
        {
           
                if (_kfs.ADD_KFS(kfsmc, kfslxr, kfslxrdh, bz))
                    return Json(SResult.Success);
                else
                    return Json(SResult.Error("新增用户失败！"));
        }
        [HttpPost]
        public JsonResult DEL_KFS(string kfsbh)
        {
            _kfs.DEL_KFS_YHFZKFS(kfsbh);
            _kfs.DEL_KFS_FZR(kfsbh);
            if (_kfs.DEL_KFS(kfsbh))
                return Json(SResult.Success);
            else
                return Json(SResult.Error("新增用户失败！"));
        }

        public JsonResult GET_KFSMBLB(string kfsbh,int pagesize,int pagenow)
        {
            var obj = _kfs.FIND_KFS_MB(kfsbh, pagesize, pagenow);
            var s = new
            {
                pagenow = obj.PageNumber,
                datacount = obj.TotalPageCount,
                d = obj
            };
            return Json(s);
        }

        public JsonResult GET_YHFZKFSLB(string tj,string yhbh,int pagesize,int pagenow)
        {
            var obj = _kfs.FIND_YHFZKFSBH(yhbh, tj, pagesize, pagenow);
            var s = new
            {
                pagenow = obj.PageNumber,
                datacount = obj.TotalPageCount,
                d = obj
            };
            return Json(s);
        }
        public JsonResult del_yhfzkfs(string yhbh,string kfsbh)
        {
            if (_kfs.DEL_KFS_YHFZKFS(yhbh, kfsbh))
                return Json(SResult.Success);
            else return Json(SResult.Error("删除授权失败"));

        }
        public JsonResult add_yhfzkfs(string yhbh, string kfsbh)
        {
            if (_kfs.ADD_KFS_YHFZKFS(yhbh, kfsbh))
                return Json(SResult.Success);
            else return Json(SResult.Error("删除授权失败"));

        }
        #endregion
        #endregion

    }
}