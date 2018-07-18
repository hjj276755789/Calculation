using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace 管理网站
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
        }

        protected void Application_OnAuthenticateRequest(Object sender, EventArgs e)
        {
            if (IsValidUrl(Context.Request.RawUrl))
            {
                IIdentity identity = Context.User == null ? new GenericIdentity(string.Empty) : Context.User.Identity;
                System.Threading.Thread.CurrentPrincipal = Context.User = new CurrentUser(identity);



                //if (identity is GenericIdentity)
                //{
                //    System.Threading.Thread.CurrentPrincipal = Context.User = CurrentUser.Debuger;
                //}
                //else
                //{
                //    System.Threading.Thread.CurrentPrincipal = Context.User = new CurrentUser(identity);
                //}             
            }

        }

        private bool IsValidUrl(string url)
        {
            return string.IsNullOrEmpty(System.IO.Path.GetExtension(url));
        }


    }
}
