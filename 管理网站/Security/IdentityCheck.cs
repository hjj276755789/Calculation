﻿using Calculation.Dal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace 管理网站
{
    public class IdentityCheck: AuthorizeAttribute
    {
        private FW_QXGL_DataProvider _qx;
        //判断是否有用户信息
        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            if (httpContext == null || httpContext.User == null)
            {
                return false;
            }
            else
            {
                var user = new CurrentUser(httpContext.User.Identity);
                if (user == null || !user.IsAuthenticated)
                {
                    return false;
                }
               
                return true;
            }
        }
        /// <summary>
        /// 判断请求来源（页面、ajax）
        /// </summary>
        /// <param name="filterContext"></param>
        protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
        {
            if (filterContext == null)
            {
                throw new ArgumentNullException();
            }

            if (!filterContext.HttpContext.Request.IsAjaxRequest())
            {
                filterContext.Result = new RedirectResult("/account/login");

            }
            else
            {
                filterContext.Result = new JavaScriptResult() { Script = "<script type='text/javascript'>top.location.href='/account/login';</script>" };
            }
        }
        public override void OnAuthorization(AuthorizationContext filterContext)
        {

            var actionName = filterContext.ActionDescriptor.ActionName;
            var controllerName = filterContext.ActionDescriptor.ControllerDescriptor.ControllerName;
            var user = new CurrentUser(filterContext.HttpContext.User.Identity);
            if (user != null&&user.IsAuthenticated)
            {
                var userid = user.YHID;
                _qx = new FW_QXGL_DataProvider();
                if (!_qx.HAS_POWER(userid, controllerName, actionName))
                {
                    filterContext.Result = new RedirectToRouteResult(new RouteValueDictionary(new { controller = "account", action = "login", returnUrl = filterContext.HttpContext.Request.Url, returnMessage = "您无权查看." }));
                    return;
                }
            }
            base.OnAuthorization(filterContext);

        }

        private bool IsAuthorized(HttpContextBase actionContext)
        {
            return true;
        }
    }
}