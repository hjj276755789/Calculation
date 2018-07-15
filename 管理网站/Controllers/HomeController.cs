﻿using Calculation.Dal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace 管理网站.Controllers
{
    public class HomeController : BaseController
    {

        FW_QXGL_DataProvider gl;

        public HomeController()
        {
            gl = new FW_QXGL_DataProvider();
        }
        /// <summary>
        /// 首页
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }
        /// <summary>
        /// 登陆页
        /// </summary>
        /// <returns></returns>
        [HttpGet]
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
        public ActionResult Login(string username,string password)
        {
            if (gl.CHECK_LOGIN(username, password))
            {
                CurrentUser.SignIn(username);
                return RedirectToAction("index");
            }
            else
            {
                this.ViewBag.message = "用户名或密码错误！";
                return View();
            }
        }

     }
}