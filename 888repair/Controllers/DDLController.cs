using _888repair.Db;
using _888repair.Models;
using _888repair.Models.Area;
using _888repair.Models.Kind;
using _888repair.Models.State;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace _888repair.Controllers
{
    [App_Start.AuthFilter]
    public class DDLController : Controller
    {
        // GET: DDL
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult getBuilding(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT distinct SystemCategory,Buliding  Building FROM[888_KsNorth].[dbo].[area] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and Buliding = @keyWord";
            }
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            var list = db.Query<AreaModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getChargePerson(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT charge_id, EmpNo, FullName ,UpdateUser,UpdateTime FROM [888_KsNorth].[dbo].[charge] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and EmpNo = @keyWord";
            }
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            var list = db.Query<DirectorModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getKind(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT kind_id KindID,SystemCategory,Sort,KindCategory,Remark FROM[888_KsNorth].[dbo].[kind] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and kind_id = @keyWord";
            }
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            sql += " order by sort asc";
            var list = db.Query<KindModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getKindBySystemCategory(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT kind_id KindID,SystemCategory,Sort,KindCategory,Remark FROM[888_KsNorth].[dbo].[kind] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and SystemCategory = @keyWord";
            }
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            sql += " order by sort asc";
            var list = db.Query<KindModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getChargePersonBySystemCategory(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT SystemCategory,charge_id, EmpNo, FullName ,UpdateUser,UpdateTime FROM [888_KsNorth].[dbo].[charge] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and SystemCategory = @keyWord";
            }
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            var list = db.Query<DirectorModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getBuildingBySystemCategory(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT distinct SystemCategory,Buliding Building FROM[888_KsNorth].[dbo].[area] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and SystemCategory = @keyWord";
            }
            //var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            //if (!string.IsNullOrEmpty(group))
            //{
            //    switch (group)
            //    {
            //        case "资讯":
            //            sql += " and SystemCategory = 'IT(资讯类)' ";
            //            break;
            //        case "后勤":
            //            sql += " and SystemCategory = 'Logistics(总务后勤类)'";
            //            break;
            //        default:
            //            break;
            //    }
            //}
            var list = db.Query<AreaModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getLoactionByBuilding(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT area_id AreaId,SystemCategory,Buliding,Location FROM[888_KsNorth].[dbo].[area] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and Buliding = @keyWord";
            }
            //var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            //if (!string.IsNullOrEmpty(group))
            //{
            //    switch (group)
            //    {
            //        case "资讯":
            //            sql += " and SystemCategory = 'IT(资讯类)' ";
            //            break;
            //        case "后勤":
            //            sql += " and SystemCategory = 'Logistics(总务后勤类)'";
            //            break;
            //        default:
            //            break;
            //    }
            //}
            var list = db.Query<AreaModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getLoaction(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT area_id AreaId,SystemCategory,Buliding,Location FROM[888_KsNorth].[dbo].[area] where 1=1 ";
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            var list = db.Query<AreaModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getKindByKindCategory(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT kind_id KindID,SystemCategory,Sort,KindCategory,Remark FROM[888_KsNorth].[dbo].[kind] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and KindCategory = @keyWord";
            }
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            sql += " order by sort asc";
            var list = db.Query<KindModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getStatus(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT distinct StatusValue,StatusText FROM [888_KsNorth].[dbo].[state] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and SystemCategory = @keyWord";
            }
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            var list = db.Query<StateModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getLocationAll(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT distinct Location FROM[888_KsNorth].[dbo].[area] where 1=1 ";
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            var list = db.Query<AreaModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }
        public ActionResult getLocationAllBySC(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT distinct Location FROM [888_KsNorth].[dbo].[area] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and SystemCategory = @keyWord";
            }
            var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
            if (!string.IsNullOrEmpty(group))
            {
                switch (group)
                {
                    case "资讯":
                        sql += " and SystemCategory = 'IT(资讯类)' ";
                        break;
                    case "后勤":
                        sql += " and SystemCategory = 'Logistics(总务后勤类)'";
                        break;
                    default:
                        break;
                }
            }
            var list = db.Query<AreaModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getKindEvery(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT kind_id KindID,SystemCategory,Sort,KindCategory,Remark FROM[888_KsNorth].[dbo].[kind] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and kind_id = @keyWord";
            }
            sql += " order by sort asc";
            var list = db.Query<KindModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }

        public ActionResult getStatusEvery(string keyWord)
        {
            RepairDb db = new RepairDb();
            string sql = "SELECT distinct StatusValue,StatusText FROM [888_KsNorth].[dbo].[state] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and SystemCategory = @keyWord";
            }
            var list = db.Query<StateModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }
    }
}