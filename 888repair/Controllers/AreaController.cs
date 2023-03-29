using _888repair.Db;
using _888repair.Models;
using _888repair.Models.Area;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace _888repair.Controllers
{
    public class AreaController : Controller
    {
        // GET: Area 辖区维护
        public ActionResult AreaIndex()
        {
            return View();
        }

        public ActionResult getAreaList(AreaModel model)
        {
            var list = new List<AreaModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT area_id AreaId,SystemCategory,Buliding,Location,
                                                  UpdateUser,UpdateTime  FROM [888_KsNorth].[dbo].[area]
                                                where 1=1");
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and SystemCategory = @SystemCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.Buliding))
                    {
                        sql += " and Buliding like '%" + model.Buliding + "%'";
                    }
                    if (!string.IsNullOrEmpty(model.Location))
                    {
                        sql += " and Location like '%" + model.Location + "%'";
                    }
                    sql += " ORDER BY area_id asc";

                    list = db.Query<AreaModel>(sql, model).ToList();
                }
                return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult editArea(AreaModel model)
        {
            try
            {
                string sql = "";
                using (RepairDb db = new RepairDb())
                {
                    if (string.IsNullOrEmpty(model.AreaId))
                    {
                        sql = string.Format(@" INSERT INTO  [888_KsNorth].[dbo].[area] (SystemCategory,Buliding,Location)VALUES(@SystemCategory,@Buliding,@Location)");
                    }
                    else
                    {
                        sql = string.Format(@" update [888_KsNorth].[dbo].[area] set SystemCategory = @SystemCategory,Buliding = @Buliding,Location=@Location where area_id = @AreaId ");
                    }
                    Dictionary<string, object> trans = new Dictionary<string, object>();
                    trans.Add(sql, model);
                    db.DoExtremeSpeedTransaction(trans);
                }
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
            return Json(new FlagTips { IsSuccess = true });
        }

        public ActionResult deleteArea(List<AreaModel> deleteList)
        {
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    var modelList = new List<AreaModel>();
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new AreaModel();
                        deleteModel.AreaId = model.AreaId;
                        deleteModel.SystemCategory = model.SystemCategory;
                        deleteModel.Buliding = model.Buliding;
                        string sql = string.Format(@" DELETE FROM [888_KsNorth].[dbo].[area]  WHERE area_id = @AreaId ");

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
    }
}