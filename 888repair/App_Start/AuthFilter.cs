using System;
using System.Collections.Generic;
using System.Linq;
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


                /*
            返回到了登陆界面，并且弹出提示信息
                */
                var url = "~/Login/LoginIndex";
                // window.location.href = url;
                filterContext.HttpContext.Response.Write("<script>window.location.href='" + url + "';</script>");
                // filterContext.HttpContext.Response.Redirect("~/Login/LoginIndex");//跳转到登陆界面
                //base.OnActionExecuting(filterContext);

                filterContext.Result = new EmptyResult();

                return;
            }

            base.OnActionExecuting(filterContext);



        }
    }
}