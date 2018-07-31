using Aspose.Cells;
using Calculation.Base;
using Calculation.Dal;
using Calculation.Models.Enums;
using Calculation.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace 管理网站.Controllers
{
    public class DataController : BaseController
    {
        private RWGL_DataProvider rw;
        private Data_DataProvider data;
        public DataController()
        {
            rw = new RWGL_DataProvider();
            data = new Data_DataProvider();
        }

        #region 数据任务
        #region 页面区
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// 添加计划
        /// </summary>
        /// <returns></returns>
        public PartialViewResult TJJH()
        {
            return PartialView();
        }
        /// <summary>
        /// 查看计划
        /// </summary>
        /// <returns></returns>
        public ActionResult CKJH()
        {
            return View();
        }
        /// <summary>
        /// 添加数据
        /// </summary>
        /// <param name="nf"></param>
        /// <param name="zc"></param>
        /// <returns></returns>
        public PartialViewResult TJSJ(int nf,int zc)
        {
            this.ViewBag.nf = nf;
            this.ViewBag.zc = zc;
            this.ViewBag.data = data.GET_JHXQ(nf, zc);
            return PartialView();
        }
        #endregion
        #region 数据区
        /// <summary>
        /// 提交计划年份
        /// </summary>
        /// <param name="nf">年份</param>
        /// <returns></returns>
        public JsonResult ADD_JH(int nf)
        {
            return Json(data.ADD_JH(nf, Base_date.GET_Z_OF_Y(nf)), JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 提交计划年份
        /// </summary>
        /// <param name="nf">年份</param>
        /// <returns></returns>
        public JsonResult get_rwxq(int nf)
        {
            return Json(data.GET_JHXQ(nf), JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult ADD_CJJL()
        {
            int nf = Int32.Parse(Request.Form["nf"]);
            int zc = Int32.Parse(Request.Form["zc"]);
            HttpPostedFileBase f = Request.Files["cjjl"];           
            Workbook workbook = new Workbook(f.InputStream);
            Cells cs = workbook.Worksheets[0].Cells;
            DataTable dt = cs.ExportDataTableAsString(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);
            if (Calculation.Dal.ZB_Data_CJBA_DataProvider.Insert(dt, nf, zc, Base_date.GET_ZCMC(nf, zc)) > 0)
                return Json(SResult.Success);
            else return Json(SResult.Error("上传文件失败，请检查EX"));
        }
        public JsonResult ADD_XZYS()
        {
            int nf = Int32.Parse(Request.Form["nf"]);
            int zc = Int32.Parse(Request.Form["zc"]);
            HttpPostedFileBase f = Request.Files["xzys"];
            Workbook workbook = new Workbook(f.InputStream);
            Cells cs = workbook.Worksheets[0].Cells;
            DataTable dt = cs.ExportDataTableAsString(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);
            if (Calculation.Dal.ZB_Data_XZYS_DataProvider.Insert(dt, nf, zc, Base_date.GET_ZCMC(nf, zc)) > 0)
                return Json(SResult.Success);
            else return Json(SResult.Error("上传文件失败，请检查EX"));
        }
        public JsonResult ADD_TDCJ()
        {
            int nf = Int32.Parse(Request.Form["nf"]);
            int zc = Int32.Parse(Request.Form["zc"]);
            HttpPostedFileBase f = Request.Files["tdcj"];
            Workbook workbook = new Workbook(f.InputStream);
            Cells cs = workbook.Worksheets[0].Cells;
            DataTable dt = cs.ExportDataTableAsString(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);
            if (Calculation.Dal.ZB_Data_TDCJ_DataProvider.Insert(dt, nf, zc) > 0)
                return Json(SResult.Success);
            else return Json(SResult.Error("上传文件失败，请检查EX"));
        }
        public JsonResult ADD_RGSJ()
        {
            int nf = Int32.Parse(Request.Form["nf"]);
            int zc = Int32.Parse(Request.Form["zc"]);
            HttpPostedFileBase f = Request.Files["rgsj"];
            Workbook workbook = new Workbook(f.InputStream);
            Cells cs = workbook.Worksheets[0].Cells;
            DataTable dt = cs.ExportDataTableAsString(1, 0, cs.MaxDataRow, cs.MaxDataColumn + 1);
            if (Calculation.Dal.ZB_Data_RGSJ_DataProvider.Insert(dt, nf, zc, Base_date.GET_ZCMC(nf,zc)) > 0)
                return Json(SResult.Success);
            else return Json(SResult.Error("上传文件失败，请检查EX"));
        }
        #endregion





        public JsonResult JHXQ(int nf)
        {
            return Json("sadf");
            //return data()
        }
        #endregion

        #region 周报

        public PartialViewResult zb_data(int rwid,int nf,int zc)
        {
            this.ViewBag.rcd = rw.GET_RWZT(rwid,nf, zc);
            this.ViewBag.data = data.GET_JHXQ(nf,zc);
            return PartialView();
        }

        public JsonResult HLSJ(int rwid,int ztlx)
        {
            switch (ztlx)
            {
                case 1: { rw.SET_DATA_ZT_CJ(rwid, DATA_ZT.确认忽略);};break;
                case 2: { rw.SET_DATA_ZT_XZ(rwid, DATA_ZT.确认忽略); }; break;
                case 3: { rw.SET_DATA_ZT_TD(rwid, DATA_ZT.确认忽略); }; break;
                case 4: { rw.SET_DATA_ZT_RG(rwid, DATA_ZT.确认忽略); }; break;
                default:return Json(false);
            };
            return Json(true);
        }
        public JsonResult SJQR(int rwid)
        {
            if (rw.SET_RWZT(rwid, RW_ZT.参数填写阶段))
                return Json(SResult.Success);
            else return Json(SResult.Error("数据确认出错"));
        }
        #endregion
    }
}