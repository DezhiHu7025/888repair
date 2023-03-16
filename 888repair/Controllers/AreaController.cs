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
                    string sql = string.Format(@"SELECT area_id AreaId,系統類別 SystemCategory,大樓別 Buliding,位置 Location,
                                                  權限 Permission,OU OU  FROM [888_north].[dbo].[area]
                                                where 1=1 and  權限 = '資訊' ");
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and 系統類別 = @SystemCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.Buliding))
                    {
                        sql += " and 大樓別 = @Buliding ";
                    }
                    if (!string.IsNullOrEmpty(model.Location))
                    {
                        sql += " and 位置 = @Location ";
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
                        sql = string.Format(@" INSERT INTO  [888_north].[dbo].[area] (系統類別,大樓別,位置,權限)VALUES(@SystemCategory,@Buliding,@Location,'資訊')");
                    }
                    else
                    {
                        sql = string.Format(@" update [888_north].[dbo].[area] set 系統類別 = @SystemCategory,大樓別 = @Buliding,位置=@Location where area_id = @AreaId and 權限 = '資訊' ");
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
                        string sql = string.Format(@" DELETE FROM [888_north].[dbo].[area]  WHERE area_id = @AreaId AND 系統類別 = @SystemCategory AND 大樓別 = @Buliding AND  權限 = '資訊' ");

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