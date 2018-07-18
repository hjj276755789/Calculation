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
                return this.User as CurrentUser;
                //if (HttpContext!=null&&HttpContext.User != null)
                //    return new CurrentUser(HttpContext.User.Identity);
                //else return null;
            }
        }
        
    }
}