using Calculation.Dal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace 管理网站.Controllers
{
    public class AccountController : Controller
    {
        FW_QXGL_DataProvider gl;
        public AccountController()
        {
            gl = new FW_QXGL_DataProvider();
        }

        [HttpGet]
        // GET: Account
        public ActionResult Login()
        {
            return View();
        }


        /// <summary>
        /// 登陆提交地址
        /// </summary>
        /// <param name="username">用户名</param>
        /// <param name="password">密码</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Login(string username, string password)
        {
            if (gl.CHECK_LOGIN(username, password))
            {
                CurrentUser.SignIn(username);
                return RedirectToAction("index","home");
            }
            else
            {
                this.ViewBag.message = "用户名或密码错误！";
                return View();
            }
        }
        public ActionResult Logout()
        {
            CurrentUser.SignOut();
            return RedirectToAction("login", "Account");
        }

        public ActionResult test()
        {
            return View();
        }
    }
}