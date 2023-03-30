using _888repair.Db;
using _888repair.Models;
using _888repair.Models.Kind;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace _888repair.Controllers
{
    public class KindController : Controller
    {
        // GET: Kind 业管维护
        public ActionResult KindIndex()
        {
            return View();
        }


        public ActionResult getKindList(KindModel model)
        {
            var list = new List<KindModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT [kind_id] KindID,SystemCategory,Sort,KindCategory,Remark,UpdateUser,UpdateTime
                                                     FROM [888_KsNorth].[dbo].[kind]
                                                where 1=1  ");
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and SystemCategory = @SystemCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.KindCategory))
                    {
                        sql += " and KindCategory like '%" + model.KindCategory + "%' ";
                    }
                    sql += " ORDER BY Sort ASC ";

                    list = db.Query<KindModel>(sql, model).ToList();
                }
                return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult editKind(KindModel model)
        {
            try
            {
                string sql = "";
                using (RepairDb db = new RepairDb())
                {
                    model.UpdateUser = "dezhi_hu";
                    model.UpdateTime = DateTime.Now;
                    if (string.IsNullOrEmpty(model.KindID))
                    {
                        sql = string.Format(@" INSERT INTO  [888_KsNorth].[dbo].[kind] ([SystemCategory],[KindCategory],[Remark],[Sort],[UpdateUser],[UpdateTime])VALUES(@SystemCategory,@KindCategory,@Remark,@Sort,@UpdateUser,@UpdateTime)");
                    }
                    else
                    {
                        sql = string.Format(@" update [888_KsNorth].[dbo].[kind] set SystemCategory = @SystemCategory,KindCategory = @KindCategory,Remark=@Remark,Sort=@Sort,UpdateUser=@UpdateUser,UpdateTime=@UpdateTime where kind_id = @KindID ");
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

        public ActionResult deleteKind(List<KindModel> deleteList)
        {
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    var modelList = new List<KindModel>();
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new KindModel();
                        deleteModel.KindID = model.KindID;
                        deleteModel.SystemCategory = model.SystemCategory;
                        deleteModel.KindCategory = model.KindCategory;
                        string sql = string.Format(@" DELETE FROM [888_KsNorth].[dbo].[kind]  WHERE kind_id = @KindID AND SystemCategory = @SystemCategory AND KindCategory = @KindCategory ");

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