using Calculation.Dal;
using Calculation.Models.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace 管理网站.Controllers
{
    public class BaseController : Controller
    {
        // GET: Base
        public CurrentUser CurrentUser
        {
            get
            {
                CurrentUser.SignIn(CurrentUser.Test.YHID);
                return CurrentUser.Test;
                //return this.User as CurrentUser;
            }
        }

        public BaseController()
        {
            if (this.CurrentUser != null)
            {
                this.ViewBag.nav = new FW_QXGL_DataProvider().GET_YHQX(this.CurrentUser.YHID);
            }
        }
        
    }
}