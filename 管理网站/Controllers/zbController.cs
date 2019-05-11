using Calculation.Base;
using Calculation.Dal;
using Calculation.Models;
using Calculation.Models.Enums;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace 管理网站.Controllers
{
    [IdentityCheck]
    public class zbController : BaseController
    {
        private FW_KFS_DataProvider kfs;
        private RWGL_DataProvider rwgl;
        private CJGL_DataProvider cjgl;
        private Param_DataProvider pagl;
        public zbController()
        {
            kfs = new FW_KFS_DataProvider();
            rwgl = new RWGL_DataProvider();
            cjgl = new CJGL_DataProvider();
            pagl = new Param_DataProvider();

        }
        // GET: zb
        #region 页面块


        /// <summary>
        /// 周报主页 ---负责开发商
        /// </summary>
        /// <returns></returns>
        
        public ActionResult index()
        {
            return View();
        }

        public ActionResult zblb(string kfsbh)
        {
            this.ViewBag.kfsbh = kfsbh;
            return View();
        }
        /// <summary>
        /// 周报任务列表
        /// </summary>
        /// <param name="mbid"></param>
        /// <param name="mbmc"></param>
        /// <param name="xflx"></param>
        /// <returns></returns>
       
        public ActionResult zb_rwlb(int mbid, string mbmc, MB_XFLX xflx)
        {
            this.ViewBag.mbid = mbid;
            this.ViewBag.mbmc = mbmc;
            this.ViewBag.xflx = xflx;
            return View();
        }
        /// <summary>
        /// 添加周报任务
        /// </summary>
        /// <param name="mbid"></param>
        /// <returns></returns>
        [IdentityCheck]
        public PartialViewResult add_zbrw(int mbid)
        {
            this.ViewBag.mbid = mbid;
            this.ViewBag.zjcs = rwgl.GET_ZJ_ZB_CS(mbid);
            this.ViewBag.bn =  DateTime.Now.Year;
            this.ViewBag.bz = Base_date.GET_Z_of_Y(DateTime.Now);
            return PartialView();
        }
        /// <summary>
        /// 添加周报参数
        /// </summary>
        /// <param name="mbid"></param>
        /// <param name="rwid"></param>
        /// <returns></returns>
        public PartialViewResult add_cs(int mbid, int rwid)
        {

            this.ViewBag.data = Param_DataProvider.GET_MBCJCSLB(mbid);
            this.ViewBag.rwid = rwid;
            return PartialView();
        }
        /// <summary>
        /// 添加码板定稿
        /// </summary>
        /// <param name="rwid"></param>
        /// <returns></returns>
        public PartialViewResult add_mbdg(int rwid)
        {
            this.ViewBag.data = rwgl.GET_RWXQ(rwid);
            return PartialView();
        }
        /// <summary>
        /// 上传竞品项目推广图片
        /// </summary>
        /// <returns></returns>
        public PartialViewResult upload_jpxmtgtp()
        {
            return PartialView();
        }
        public PartialViewResult sczb(int mbid,int nf,int zc)
        {
            this.ViewBag.serverpath = ConfigurationManager.AppSettings["SerPath"];
            this.ViewBag.serverpoint = ConfigurationManager.AppSettings["SerPoint"];
            this.ViewBag.mbid = mbid;
            this.ViewBag.nf = nf;
            this.ViewBag.zc = zc;
            return PartialView();
        }
        #endregion

        #region 数据接口块

        ///获取开发商列表
        public JsonResult get_kfslb(string tj,int pagesize,int pagenow)
        {
            var obj = kfs.FIND_YHFZKFS(this.CurrentUser.YHBH, tj, pagesize, pagenow);
            var s = new
            {
                pagenow = obj.PageNumber,
                datacount = obj.TotalPageCount,
                d = obj
            };
            return Json(s);
        }

        ///获取周报模板类型
        [HttpPost]
       
        public JsonResult get_zbmblx(int pagesize, int pagenow,string mbmc)
        {
            return Json(rwgl.GET_ZB_LB(this.CurrentUser.YHBH,pagesize, pagenow,mbmc));
        }
        /// <summary>
        /// 获取后保任务列表
        /// </summary>
        /// <param name="mbid">模板ID</param>
        /// <param name="pagesize">分页大小</param>
        /// <param name="pagenow">分页页码</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult get_zbrwlb(int mbid, int pagesize, int pagenow)
        {
            var data = rwgl.GET_ZB_RWLB(mbid, pagesize, pagenow).ToList();
            return Json(data);
        }
        /// <summary>
        /// 添加任务
        /// </summary>
        /// <param name="rwmc">任务名称</param>
        /// <param name="mbid">模板ID</param>
        /// <param name="nf">年份</param>
        /// <param name="zc">周次</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult add_zbrw(string rwmc, int mbid, int nf, int zc,int zjcs_nf,int zjcs_zc)
        {
            if (rwgl.Add_ZB(rwmc, mbid, nf, zc))
            {
                var t = rwgl.GET_RWXQ(mbid,nf,zc);
                if (Param_DataProvider.jcszsz(t.rwid, mbid, zjcs_nf, zjcs_zc).IsSuccessful)
                {
                    rwgl.SET_RWZT(t.rwid, RW_ZT.文档生成中);
                }
                return Json(SResult.Success);
            }
            else
            {
                return Json(SResult.Error("发布任务失败"));
            }

        }
        /// <summary>
        /// 生成PPT接口
        /// </summary>
        /// <param name="mbid"></param>
        /// <param name="nf"></param>
        /// <param name="zc"></param>
        /// <returns></returns>
        public JsonResult sc(int mbid, int nf, int zc)
        {
            string url = ConfigurationManager.AppSettings["SerPath"] + "?mbid=" + mbid + "&nf=" + nf + "&zc=" + zc;
            try
            {
                string sql = HttpHelper.GetResponseString(HttpHelper.HttpPost(url, null, 3000));

                return Json(SResult.Success);
            }
            catch (Exception)
            {

                return Json( SResult.Error("生成报表服务未启动或服务器无法链接"));
            }
            
        }
        /// <summary>
        /// 添加文字参数
        /// </summary>
        /// <param name="rwid">任务id</param>
        /// <param name="csid">参数id</param>
        /// <param name="csnr">参数内容</param>
        /// <param name="sfbl">是否并列</param>
        /// <returns></returns>
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
        /// <summary>
        /// 获取文字参数
        /// </summary>
        /// <param name="rwid"></param>
        /// <param name="csid"></param>
        /// <returns></returns>
        public JsonResult get_wzcs(int rwid, int csid)
        {
            return Json(Param_DataProvider.GET_RWCSNR(rwid, csid));
        }
        /// <summary>
        /// 删除参数
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public JsonResult del_csnr(int id)
        {
            return Json(Param_DataProvider.DEL_RWCJCS(id) ? SResult.Success : SResult.Error("删除失败"));
        }

        /// <summary>
        /// 通过参数状态
        /// </summary>
        /// <param name="rwid">任务ID</param>
        /// <returns></returns>
        public JsonResult tgcssz(int rwid)
        {
            return Json(rwgl.SET_RWZT(rwid, RW_ZT.文档生成中) ? SResult.Success : SResult.Error("设置失败"));
        }


        /// <summary>
        /// 添加文件参数
        /// </summary>
        /// <returns></returns>
        public JsonResult add_wjcs()
        {
            int nf = Request.Form["nf"].ints();
            int zc = Request.Form["zc"].ints();
            string cjmc = Request.Form["cjmc"];
            int rwid = Request.Form["rwid"].ints();
            int csid = Request.Form["csid"].ints();
            HttpPostedFileBase f = Request.Files["rgsj"];
            string path = ConfigurationManager.AppSettings["ParamPath"] + nf + "\\" + zc + "\\" + cjmc;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string abpath = Path.Combine(path, Base_IdHelper.GetID() + ".pptx");
            f.SaveAs(abpath);
            int id = Param_DataProvider.RESET_RWCJCS(rwid, csid, abpath);
            if (id != -1)
            {
                SResult s = SResult.Success;
                s.Data = id;
                return Json(s);
            }
            return Json(SResult.Error("添加失败"));
        }
        /// <summary>
        /// 删除任务
        /// </summary>
        /// <param name="rwid"></param>
        /// <returns></returns>
        public JsonResult del_rw(int rwid)
        {
            if (rwgl.Del_ZB(rwid))
                return Json(SResult.Success);
            else return Json(SResult.Error("删除失败"));
        }



        /// <summary>
        /// 添加定稿文件
        /// </summary>
        /// <returns></returns>
        public JsonResult add_dgwj()
        {
            int nf = Request.Form["nf"].ints();
            int zc = Request.Form["zc"].ints();
            int rwid = Request.Form["rwid"].ints();
            HttpPostedFileBase f = Request.Files["dgwj"];
            string path = ConfigurationManager.AppSettings["ParamPath"] + nf + "\\" + zc + "\\" + rwid;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string abpath = Path.Combine(path, Base_IdHelper.GetID() + ".pptx");
            f.SaveAs(abpath);
            if (Param_DataProvider.SET_RWDGWJ(rwid, abpath))
            {
                SResult s = SResult.Success;
                return Json(s);
            }
            return Json(SResult.Error("添加失败"));
        }
        /// <summary>
        /// 上传推广图片
        /// </summary>
        /// <returns></returns>
        public JsonResult add_tgtp()
        {
            int nf = Request.Form["nf"].ints();
            int zc = Request.Form["zc"].ints();
           
            HttpPostedFileBase f = Request.Files["tgtp"];
            string path = ConfigurationManager.AppSettings["DgPath"] + nf + "\\" + zc;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string abpath = Path.Combine(path, "bak.zip");
            f.SaveAs(abpath);
            using (ZipArchive archive = System.IO.Compression.ZipFile.Open(abpath, ZipArchiveMode.Update))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)|| entry.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                    {
                        entry.ExtractToFile(Path.Combine(path, entry.Name),true);
                    }
                }

            }
            return Json(SResult.Success);
        }
        #endregion

        #region 导出模块
        public FileStreamResult export(int rwid)
        {
            Rw_Item_Model rim = rwgl.GET_RWXQ(rwid);

            return File(new FileStream(rim.xzdz, FileMode.Open), "application/octet-stream", Url.Encode(rim.rwmc + ".pptx"));
        }
        #region 导出模块
        public FileStreamResult export_dg(int rwid)
        {
            Rw_Item_Model rim = rwgl.GET_RWXQ(rwid);

            return File(new FileStream(rim.xzdz2, FileMode.Open), "application/octet-stream", Url.Encode(rim.rwmc + ".pptx"));
        }
        #endregion
        #endregion
    }

    /// <summary>
    /// 周报—竞品
    /// </summary>
    public class jp_zbController : BaseController {
        private RWGL_DataProvider rwgl;
        #region 页面
        public ActionResult Index(int rwid,int mbid)
        {
            RWGL_DataProvider rw = new RWGL_DataProvider();
            rwgl = new RWGL_DataProvider();
            this.ViewBag.rwxq = rw.GET_RWXQ(rwid);
            this.ViewBag.mbid = mbid;
            return View();
        }
        /// <summary>
        /// 本案竞争格局
        /// </summary>
        /// <returns></returns>
        public ActionResult Bajzgj(int id, int nf, int zc)
        {
            this.ViewBag.nf = nf;
            this.ViewBag.zc = zc;
            this.ViewBag.id = id;
            return View();
        }
        /// <summary>
        /// 竞品竞争格局
        /// </summary>
        /// <returns></returns>
        public ActionResult Set_Jpjzgj_Param(int id, int nf, int zc)
        {
            this.ViewBag.nf = nf;
            this.ViewBag.zc = zc;
            this.ViewBag.id = id;
            return View();
        }
        /// <summary>
        /// 管理竞品项目
        /// </summary>
        /// <param name="baid"></param>
        /// <returns></returns>
        public ActionResult Jpxm(int baid, int nf, int zc)
        {
            this.ViewBag.baid = baid;
            this.ViewBag.nf = nf;
            this.ViewBag.zc = zc;
            return View();
        }
        /// <summary>
        /// 选取竞争格局范围
        /// </summary>
        public PartialViewResult Jzgjfw(int baid)
        {
            this.ViewBag.baid = baid;
            this.ViewBag.jzgjlb = ZB_Param_JP_DataProvider.GET_JPGJ();
            return PartialView();
        }

        public PartialViewResult Jcjzcs(int mbid,int rwid)
        {
            this.ViewBag.mbid = mbid;
            this.ViewBag.rwid = rwid;
            return PartialView();
        }
        #endregion

        #region 数据

        /// <summary>
        /// 通用成交备案数据查询接口
        /// </summary>
        /// <param name="nf"></param>
        /// <param name="zc"></param>
        /// <param name="zt"></param>
        /// <param name="qy"></param>
        /// <param name="lpmc"></param>
        /// <param name="yt"></param>
        /// <param name="xfyt"></param>
        /// <param name="hx"></param>
        /// <param name="pagesize"></param>
        /// <param name="pagenow"></param>
        /// <returns></returns>
        public JsonResult cxjg(int nf, int zc, string[] zt, string[] qy, string[] kfs, string[] lpmc, string[] yt, string[] xfyt, string[] hx, int? pagesize, int? pagenow)
        {
            IPageList<Data_Cjba_Default> list = null;
            if (zt != null || qy != null ||kfs!=null|| lpmc != null || yt != null || xfyt != null || hx != null) {
                JP_ParamValueModel param = new JP_ParamValueModel();
                param.zt = zt;
                param.qy = qy;
                param.kfs = kfs;
                param.lpmc = lpmc;
                param.yt = yt;
                param.xfyt = xfyt;
                param.hx = hx;

                
                list = Param_DataProvider.GET_JP_CJBAXX(nf, zc, param, pagesize.HasValue ? pagesize.Value : 10, pagenow.HasValue ? pagenow.Value : 1);
            }
            else
            {
                list = Param_DataProvider.GET_JP_CJBAXX(nf, zc, pagesize.HasValue ? pagesize.Value : 10, pagenow.HasValue ? pagenow.Value : 1);
            }
            var s = new
            {
                pagenow = list.PageNumber,
                datacount = list.TotalPageCount,
                data = list
            };
            return Json(s);


        }

        /// <summary>
        /// 获取竞品本案
        /// </summary>
        /// <param name="rwid"></param>
        /// <returns></returns>
        public JsonResult get_ba(int rwid)
        {
            return Json(Param_DataProvider.GET_JP_BA(rwid));
        }
        public JsonResult add_ba(int rwid, string bamc)
        {
            if (Param_DataProvider.ADD_JP_BA(rwid, bamc))
                return Json(SResult.Success);
            else return Json(SResult.Error("添加失败"));
        }
        public JsonResult del_ba(int id)
        {
            if (Param_DataProvider.DEL_JP_BA(id))
                return Json(SResult.Success);
            else return Json(SResult.Error("删除失败"));
        }
        public JsonResult save_baxmcs(string[] zt, string[] qy, string[] kfs,string[] lpmc, string[] yt, string[] xfyt, string[] hx, string[] zlmjqj,string qtcs,int id)
        {
            if (zt != null || qy != null ||kfs!=null|| lpmc != null || yt != null || xfyt != null || hx != null|| zlmjqj!=null||qtcs!=null)
            {
                JP_ParamValueModel param = new JP_ParamValueModel();
                param.zt = zt;
                param.qy = qy;
                param.kfs = kfs;
                param.lpmc = lpmc;
                param.yt = yt;
                param.xfyt = xfyt;
                param.hx = hx;
                param.zlmjqj = zlmjqj;
                param.qtcs = qtcs;
                if (Param_DataProvider.SAVE_JP_BAXMCS(id, param))
                    return Json(SResult.Success);
                else return Json(SResult.Error("保存失败"));
            }
            else
            { return Json(SResult.Error("竞品参数为空")); }

        }
        public JsonResult get_jpba_xq(int id)
        {
            var T = Param_DataProvider.GET_JP_BA_XQ(id);
            return Json(T);
        }

        /// <summary>
        /// 获取竞品项目
        /// </summary>
        /// <param name="baid"></param>
        /// <returns></returns>
        public JsonResult get_jpxm(int baid)
        {
            var T = Param_DataProvider.GET_JP_JPXM(baid);
            return Json(T);
        }
        public JsonResult add_jpxm(int baid, int jzgjid)
        {
            if (Param_DataProvider.add_jp_jpxm(baid, jzgjid))
                return Json(SResult.Success);
            else return Json(SResult.Error("添加失败"));

        }
        public JsonResult del_jpxm(int id)
        {
            if(Param_DataProvider.del_jp_jpxm(id))
            return Json(SResult.Success);
            else return Json(SResult.Error("删除失败"));
        }
        public JsonResult save_jpxmcs(string[] zt, string[] qy, string [] kfs,string[] lpmc, string[] yt, string[] xfyt, string[] hx,string [] zlmjqj, int id)
        {
            if (zt != null || qy != null || kfs != null || lpmc != null || yt != null || xfyt != null || hx != null|| zlmjqj!=null)
            {
                JP_ParamValueModel param = new JP_ParamValueModel();
                param.zt = zt;
                param.qy = qy;
                param.lpmc = lpmc;
                param.kfs = kfs;
                param.yt = yt;
                param.xfyt = xfyt;
                param.hx = hx;
                param.zlmjqj = zlmjqj;
                if (Param_DataProvider.SAVE_JP_JPXMCS(id, param))
                    return Json(SResult.Success);
                else return Json(SResult.Error("保存失败"));
            }
            else
            { return Json(SResult.Error("竞品参数为空")); }

        }
        public JsonResult get_jpxm_xq(int id)
        {
            var T = Param_DataProvider.GET_JP_JPXM_XQ(id);
            return Json(T);
        }

        /// <summary>
        /// 继承上周设置
        /// </summary>
        /// <param name="mbid"></param>
        /// <param name="nf"></param>
        /// <param name="zc"></param>
        /// <returns></returns>
        public JsonResult jcszsz(int mbid,int rwid,int nf,int zc)
        {
            return Json(Param_DataProvider.jcszsz(rwid, mbid, nf, zc));


        }


        #region 参数

       
        public JsonResult GET_ZTLB(string ztmc)
        {
            var list = Param_DataProvider.GET_ZTMC(ztmc);
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        [HttpGet]
        public JsonResult GET_QYLB(string qymc)
        {
            var list = Param_DataProvider.GET_QYMC(qymc);
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        [HttpGet]
        public JsonResult GET_KFSLB(string kfs)
        {
            var list = Param_DataProvider.GET_KFS(kfs);
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GET_LPLB(string lpmc)
        {
            var list = Param_DataProvider.GET_LPMC(lpmc);
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GET_YTLB(string ytmc)
        {
            var list = Param_DataProvider.GET_YTMC(ytmc);
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GET_XFYTLB(string xfyt)
        {
            var list = Param_DataProvider.GET_XFYTMC(xfyt);
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GET_HXLB(string hxmc)
        {
            var list = Param_DataProvider.GET_HXMC(hxmc);
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        #endregion
    
        #endregion



    }       
}