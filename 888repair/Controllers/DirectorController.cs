﻿using _888repair.Db;
using _888repair.Models;
using _888repair.Models.Director;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace _888repair.Controllers
{
    public class DirectorController : Controller
    {
        // GET: Director
        public ActionResult DirectorIndex()
        {
            return View();
        }

        public ActionResult getDirectorList(DirectorModel model)
        {
            var list = new List<DirectorModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT charge_id, EmpNo, FullName,SystemCategory ,UpdateUser,UpdateTime FROM [888_KsNorth].[dbo].[charge]
                                                where 1=1  ");
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and SystemCategory = @SystemCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.FullName))
                    {
                        sql += " and FullName like '%" + model.FullName + "%'";
                    }
                    sql += " ORDER BY charge_id asc";

                    list = db.Query<DirectorModel>(sql, model).ToList();
                }
                return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult editDirector(DirectorModel model)
        {
            try
            {
                model.UpdateUser = Session["fullname"].ToString();
                model.UpdateTime = DateTime.Now;
                using (RepairDb db = new RepairDb())
                {
                    string checkSql = @"select * from [888_KsNorth].[dbo].[charge] where EmpNo = @EmpNo and SystemCategory = @SystemCategory";
                    var list = db.Query<DirectorModel>(checkSql, model).ToList();
                    if(list.Count() != 0)
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "请勿重复新增" });
                    }
                    else
                    {
                        string sql = string.Format(@" INSERT INTO  [888_KsNorth].[dbo].[charge] (SystemCategory,EmpNo,FullName,UpdateUser,UpdateTime)VALUES(@SystemCategory,@EmpNo,@FullName,@UpdateUser,@UpdateTime)");
                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(sql, model);
                        db.DoExtremeSpeedTransaction(trans);
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
            return Json(new FlagTips { IsSuccess = true });
        }

        public ActionResult deleteDirector(List<DirectorModel> deleteList)
        {
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    var modelList = new List<DirectorModel>();
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new DirectorModel();
                        deleteModel.charge_id = model.charge_id;
                        deleteModel.SystemCategory = model.SystemCategory;
                        deleteModel.EmpNo = model.EmpNo;
                        string sql = string.Format(@" DELETE FROM [888_KsNorth].[dbo].[charge]  WHERE charge_id = @charge_id AND EmpNo = @EmpNo and SystemCategory = @SystemCategory");

                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(sql, deleteModel);
                        db.DoExtremeSpeedTransaction(trans);
                    }

                }
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
            return Json(new FlagTips { IsSuccess = true });
        }

        /// <summary>
        /// 获取员工姓名
        /// </summary>
        /// <param name="AccountName"></param>
        /// <returns></returns>
        public ActionResult getEmpInfo(string EmpNo)
        {
            UserModel user = new UserModel();
            if (string.IsNullOrEmpty(EmpNo))
            {
                return Json(new FlagTips { IsSuccess = false });
            }
            try
            {
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
	   b.DeptName,
	   c.groupid,
	   c.groupname,a.password2
FROM [Common].[dbo].[kcis_account] a
LEFT JOIN [Common].[dbo].[AFS_Dept] B
ON a.deptid2 =b.DeptID_eip  
LEFT JOIN [db_forminf].[dbo].[UserGroup] c
ON a.AccountID = c.account
where sourcetype='A' and status = 'Y'
and ( a.Empno = @EmpNo)";
                    user = db.Query<UserModel>(userSql, new { EmpNo }).FirstOrDefault();
                    if(user == null)
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "未找到账号，请重新输入" });
                    }

                }
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
            return Json(user, JsonRequestBehavior.AllowGet);
        }
    }
}