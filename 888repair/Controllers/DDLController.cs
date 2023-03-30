using _888repair.Db;
using _888repair.Models;
using _888repair.Models.Area;
using _888repair.Models.Kind;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace _888repair.Controllers
{
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
            string sql = "SELECT area_id AreaId,SystemCategory,Buliding,Location FROM[888_KsNorth].[dbo].[area] where 1=1 ";
            if (!string.IsNullOrEmpty(keyWord))
            {
                sql += " and area_id = @keyWord";
            }
            sql += " order by sortno asc";
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
            sql += " order by sort asc";
            var list = db.Query<KindModel>(sql, new { keyWord }).ToList();
            return Json(list, JsonRequestBehavior.AllowGet);
        }
    }
}