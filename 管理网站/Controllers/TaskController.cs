using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace 管理网站.Controllers
{
    public class TaskController : BaseController
    {
        // GET: Task
        public PartialViewResult Index()
        {
            return PartialView();
        }
    }
}