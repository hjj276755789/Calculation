using Calculation.Base;
using Calculation.Dal;
using Calculation.Models;
using Calculation.Models.Enums;
using Calculation.Models.Models;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace 管理网站.Controllers
{
    public class zbController : Controller
    {
        private RWGL_DataProvider rwgl;
        private CJGL_DataProvider cjgl;
        private Param_DataProvider pagl;
        public zbController()
        {
            rwgl = new RWGL_DataProvider();
            cjgl = new CJGL_DataProvider();
            pagl = new Param_DataProvider();
        }
        // GET: zb
        #region 页面块

       
        public ActionResult index()
        {
            return View();
        }
        
        public ActionResult zb_rwlb(int mbbh,string mbmc)
        {
            this.ViewBag.mbbh = mbbh;
            this.ViewBag.mbmc = mbmc;
            return View();
        }

        public PartialViewResult add_zbrw(int mbbh)
        {
            this.ViewBag.mbbh = mbbh;
            return PartialView();
        }

        public PartialViewResult add_cs(int mbbh,int rwid)
        {

            this.ViewBag.data = Param_DataProvider.GET_MBCJCSLB(mbbh);
            this.ViewBag.rwid = rwid;
            return PartialView();
        }

        #endregion

        #region 数据接口块


        [HttpPost]
        public JsonResult get_zbmblx(int pagesize, int pagenow)
        {
            return Json(rwgl.GET_ZB_LB(pagesize, pagenow));
        }

        [HttpPost]
        public JsonResult get_zbrwlb(int mbbh, int pagesize, int pagenow)
        {
            var data = rwgl.GET_ZB_RWLB(mbbh, pagesize, pagenow).ToList();
            return Json(data);
        }

        [HttpPost]
        public JsonResult add_zbrw(string rwmc,int mbbh,int nf,int zc)
        {
            if(rwgl.Add_ZB(rwmc, mbbh, nf, zc))
            return Json(SResult.Success);
            else
            {
                return Json(SResult.Error("发布任务失败"));
            }

        }

        public JsonResult sc(int mbid,int nf,int zc)
        {
            string url = ConfigurationManager.AppSettings["SerPath"]+ "?mbid="+ mbid+"&nf="+nf+"&zc="+zc;
            string sql = HttpHelper.GetResponseString( HttpHelper.HttpPost(url, null, 3000));
            return Json(sql,JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult add_wzcs(int rwid, int csid, string csnr, int sfbl)
        {
            int id = 0;
            if (sfbl == 0)
                id = Param_DataProvider.RESET_RWCJCS(rwid, csid, csnr);
            else id = Param_DataProvider.SET_RWCJCS(rwid, csid, csnr);
            if (id != -1) {
                SResult s = SResult.Success;
                s.Data = id;
                return Json(s);
            }
            return Json(SResult.Error("添加失败"));

        }
        public JsonResult del_csnr(int id)
        {
            return Json(Param_DataProvider.DEL_RWCJCS(id) ? SResult.Success : SResult.Error("删除失败"));
        }

        public JsonResult tgcssz(int rwid)
        {
            return Json(rwgl.SET_RWZT(rwid, RW_ZT.文档生成中) ? SResult.Success : SResult.Error("设置失败"));
        }

        public JsonResult add_wjcs()
        {
            int nf = Request.Form["nf"].ints();
            int zc = Request.Form["zc"].ints();
            string cjmc = Request.Form["cjmc"];
            int rwid= Request.Form["rwid"].ints();
            int csid = Request.Form["csid"].ints();
            HttpPostedFileBase f = Request.Files["rgsj"];
            string path = ConfigurationManager.AppSettings["ParamPath"]+nf+"/"+zc+"/"+ cjmc;
            if(!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            f.SaveAs(Path.Combine( path ,f.FileName));
            

            int id = Param_DataProvider.RESET_RWCJCS(rwid, csid, Path.Combine(path, f.FileName));
            if (id != -1)
            {
                SResult s = SResult.Success;
                s.Data = id;
                return Json(s);
            }
            return Json(SResult.Error("添加失败"));
        }
        #endregion

        #region 导出模块
        public FileStreamResult export(int rwid)
        {
            Rw_Item_Model rim = rwgl.GET_RWXQ(rwid);
            
            return File(new FileStream(rim.xzdz, FileMode.Open),  "application/octet-stream", Url.Encode(rim.rwmc+".pptx"));
        }
        #endregion
    }
}