using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.SessionState;

namespace _888repair.App_Start
{
    public class AuthFilter : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            // 获取当前请求的控制器和操作方法
            string controllerName = filterContext.RouteData.Values["controller"].ToString();
            string actionName = filterContext.RouteData.Values["action"].ToString();

            // 如果当前请求的是登录操作，则不进行重定向
            if (controllerName == "Login" && actionName == "doLogin")
            {
                base.OnActionExecuting(filterContext);
                return;
            }

            // 检查用户是否已登录
            if (filterContext.HttpContext.Session["fullname"] == null)
            {
                //string templatePath = string.Format("~\\Excel\\Adminission.xlsx");
                //string serverDir = System.IO.Path.Combine(System.Web.HttpContext.Current.Server.MapPath(templatePath));

                //string[] resultString = serverDir.Split(new char[] { '\\' });
                ////string str = System.Web.Hosting.HostingEnvironment.SiteName;
                //string str = resultString[2];
                //if (str == "888repair")
                //{
                //    str = null;
                //}
                //var url = str + "/Login/LoginIndex";
                // window.location.href = url;
                var url ="../Login/LoginIndex";
                filterContext.HttpContext.Response.Write("<script> window.location.href='" + url + "';</script>");
                base.OnActionExecuting(filterContext);
                filterContext.Result = new EmptyResult();
                return;
            }

            base.OnActionExecuting(filterContext);



        }

    }
}