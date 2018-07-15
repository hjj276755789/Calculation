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
        public string YHID { get; private set; }

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
                this.YHID = userData[0];
            }
        }

        public CurrentUser(string YHID)
        {
            this._identity = new System.Security.Principal.GenericIdentity(YHID);
            this.YHID = YHID;
        }

        public static void SignIn(string YHID)
        {
            DateTime now = DateTime.Now;

            var ticket = new FormsAuthenticationTicket(
                1,
                Guid.NewGuid().ToString(),
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



        IIdentity IPrincipal.Identity
        {
            get { return _identity; }
        }

        bool IPrincipal.IsInRole(string role)
        {
            throw new NotImplementedException();
        }
    }
}