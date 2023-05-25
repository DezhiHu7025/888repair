using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using _888repair.Db;
using _888repair.Models;
using _888repair.Models.Director;

namespace _888repair.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        public ActionResult LoginIndex()
        {
            return View();
        }

        /// <summary>
        /// 登录
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult doLogin(UserModel model)
        {
            if (string.IsNullOrEmpty(model.Account) || string.IsNullOrEmpty(model.Password))
            {
                string errorMsg = "【登录失败】「帐号」或[密码]栏未输入！";
                return Json(new FlagTips { IsSuccess = false, Msg = errorMsg });
            }
            else
            {
                try
                {
                    //获取登入账号信息
                    LoginController con = new LoginController();
                    UserModel user = con.getEmpInfo(model.Account);
                    if (user == null)
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "账号异常" });
                    }
                    model.fullname = user.fullname;
                    model.DeptName = user.DeptName;
                    model.Cname = user.Cname;
                    model.GroupName = user.GroupName;
                    model.Manager = user.Manager;
                    model.Viewer = user.Viewer;
                    model.EmpNo = user.EmpNo;

                    LoginController login = new LoginController();
                    if (login.checkAccount("192.168.80.222", model.Account, model.Password))
                    {
                        Session["fullname"] = model.fullname;
                        Session["GroupName"] = model.GroupName;
                        Session["Manager"] = model.Manager;
                        Session["Viewer"] = model.Viewer;
                        Session["EmpNo"] = model.EmpNo;
                    }
                    else
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "登录失败，账号或密码不正确。Login failed. The account or password is incorrect." });
                    }
                }
                catch (Exception ex)
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
                }
                finally
                {

                }
                return Json(new FlagTips { IsSuccess = true, Msg = "" });
            }
        }

        public Boolean checkAccount(string LDAP, string EmployeeID, string EmployeePasswd)
        {
            string LDAP_STRING = "LDAP://" + LDAP;

            try
            {
                using (DirectoryEntry entry = new DirectoryEntry(LDAP_STRING, EmployeeID, EmployeePasswd))
                {
                    object nativeObject = entry.NativeObject;
                    return true;
                }
            }
            catch (DirectoryServicesCOMException)
            {
                return false;
            }
        }

        public UserModel getEmpInfo(string Account)
        {
            UserModel user = new UserModel();
            using (RepairDb db = new RepairDb())
            {
                string userSql = @"SELECT a.AccountID Account,
       a.ename,
       a.cname,
       a.fullname,
       a.titlename,
       a.status,
       a.deptid2,
       a.sourcetype,
       a.email,
       B.DeptName,
       a.password password2,
       c.Permission GroupName,
	   c.Viewer,
	   c.Manager,a.EmpNo
FROM [Common].[dbo].[kcis_account] a
    LEFT JOIN [Common].[dbo].[AFS_Dept] B
        ON a.deptid2 = B.DeptID_eip
    LEFT JOIN [888_KsNorth].[dbo].[admin] c
        ON a.EmpNo = c.EmpNo
where a.AccountID = @Account";
                user = db.Query<UserModel>(userSql, new { Account }).FirstOrDefault();

            }
            return user;
        }


        public ActionResult LoginOut(UserModel model)
        {
            //if (Session["fullname"] != null)//先判断session是否存在 （sessionname表示你session的名称）
            //{
            //    Session.Abandon();
            //}
            Session.Abandon();
            return Json(new FlagTips { IsSuccess = true, Msg = "" });
        }

        public ActionResult checkDirector(string EmpNo)
        {
            string chekMum = null;
            using (RepairDb db = new RepairDb())
            {
                string checkSql = @"SELECT * FROM [888_KsNorth].[dbo].[charge] WHERE EmpNo = @EmpNo ";
                var user = db.Query<DirectorModel>(checkSql, new { EmpNo }).ToList();
                if (user.Count() > 0)
                {
                    chekMum = "1";
                }
                else
                {
                    chekMum = "0";
                }
            }
            return Json(new FlagTips { IsSuccess = true, Msg = chekMum });
        }

        /// <summary>
        /// 管理员登录
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult adminLogin(UserModel model)
        {
            if (string.IsNullOrEmpty(model.Account) || string.IsNullOrEmpty(model.Password))
            {
                string errorMsg = "【登录失败】「帐号」或[密码]栏未输入！";
                return Json(new FlagTips { IsSuccess = false, Msg = errorMsg });
            }
            else
            {
                try
                {
                    //获取登入账号信息
                    LoginController con = new LoginController();
                    UserModel user = con.getEmpInfo(model.Account);
                    if (user == null)
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "账号异常" });
                    }
                    model.fullname = user.fullname;
                    model.DeptName = user.DeptName;
                    model.Cname = user.Cname;
                    model.GroupName = user.GroupName;
                    model.Manager = user.Manager;
                    model.Viewer = user.Viewer;
                    model.EmpNo = user.EmpNo;

                    if (model.Password == "admin123456")
                    {
                        Session["fullname"] = model.fullname;
                        Session["GroupName"] = model.GroupName;
                        Session["Manager"] = model.Manager;
                        Session["Viewer"] = model.Viewer;
                        Session["EmpNo"] = model.EmpNo;
                    }
                    else
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "登录失败，账号或密码不正确。Login failed. The account or password is incorrect." });
                    }
                }
                catch (Exception ex)
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
                }
                finally
                {

                }
                return Json(new FlagTips { IsSuccess = true, Msg = "" });
            }
        }
    }
}