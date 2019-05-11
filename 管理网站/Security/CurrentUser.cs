using Calculation.Base;
using Calculation.Dal;
using Calculation.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Web;
using System.Web.Security;

namespace 管理网站
{
    public class CurrentUser : IPrincipal
    {
        private IIdentity _identity;

        /// <summary>用户编号</summary>
        public string YHBH { get; private set; }

        public QXXX activityPath { get; set; }

        /// <summary>
        /// 获有一个值，该值表示是否验证了用户
        /// </summary>
        public bool IsAuthenticated
        {
            get { return this._identity.IsAuthenticated; }
        }

        public CurrentUser(IIdentity identity)
        {
            _identity = identity;
            if (identity is FormsIdentity)
            {
                string[] userData = ((FormsIdentity)identity).Ticket.UserData.Split(',');
                this.YHBH = userData[0];
            }
        }

        public CurrentUser(string YHID)
        {
            this._identity = new System.Security.Principal.GenericIdentity(YHID);
            this.YHBH = YHID;
        }

        public static void SignIn(string YHID)
        {
            DateTime now = DateTime.Now;
            var ticket = new FormsAuthenticationTicket(
                1,
                YHID,
                now,
                now.AddDays(1),
                true,
                string.Format("{0}", YHID),
                FormsAuthentication.FormsCookiePath
                );
            var encryptedTicket = FormsAuthentication.Encrypt(ticket);
            var cookie = new HttpCookie(FormsAuthentication.FormsCookieName, encryptedTicket);
            cookie.HttpOnly = true;
            cookie.Secure = FormsAuthentication.RequireSSL;
            cookie.Path = FormsAuthentication.FormsCookiePath;

            if (FormsAuthentication.CookieDomain != null)
            {
                cookie.Domain = FormsAuthentication.CookieDomain;
            }

            HttpContext.Current.Response.Cookies.Add(cookie);

        }
        public static List<QXXX> GETPower(string YHID)
        {
            var list = new FW_QXGL_DataProvider().GET_YHQX(YHID);
            var lv1 = list.Where(m => m.fqxbh.IsNull());
            foreach (var i_1 in lv1)
            {
                i_1.GetChildNode(list);
            }
            return lv1.OrderBy(m=>m.qxbh).ToList();
        }

        public static void SignOut()
        {
            if (HttpContext.Current.User.Identity.IsAuthenticated)
            {
                FormsAuthentication.SignOut();
            }
        }

        public static CurrentUser Test
        {
            get
            {
                return new CurrentUser("admin");
            }
        }
        public  CurrentUser UserInfo
        {
            get
            {
                return this;
            }
        }


        IIdentity IPrincipal.Identity
        {
            get { return _identity; }
        }

        bool IPrincipal.IsInRole(string role)
        {
            throw new NotImplementedException();
        }

        public static QXXX IniNav(string ctr,string act)
        {
            return  new FW_QXGL_DataProvider().GET_YHQX(ctr, act);
        }
    }

    #region 权限扩展方法
    public static class EXTENDS_QXGL
    {
        /// <summary>
        ///  扩展方法 生成下级权限
        /// </summary>
        /// <param name="target">当前节点</param>
        /// <param name="list">权限目录</param>
        /// <returns></returns>
        public static QXXX GetChildNode(this QXXX target, List<QXXX> list)
        {
            List<QXXX> temp = new List<QXXX>();
            var tree = list.Where(m => m.fqxbh == target.qxbh);
            target.xjqx = new List<QXXX>();
            target.xjqx.AddRange(tree);
            return target;
        }
    }

    #endregion
}