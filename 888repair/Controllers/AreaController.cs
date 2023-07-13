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
    [App_Start.AuthFilter]
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
                    string sql = string.Format(@"SELECT area_id AreaId,SystemCategory,Buliding Building,Location,
                                                  UpdateUser,UpdateTime  FROM [888_KsSouth].[dbo].[area]
                                                where 1=1");
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and SystemCategory = @SystemCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.Building))
                    {
                        sql += " and Buliding like '%" + model.Building + "%'";
                    }
                    if (!string.IsNullOrEmpty(model.Location))
                    {
                        sql += " and Location like '%" + model.Location + "%'";
                    }
                    var group = Session["GroupName"].ToString();
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
                    sql += " ORDER BY sortno asc, area_id asc";

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
                model.UpdateUser = Session["fullname"].ToString();
                model.UpdateTime = DateTime.Now;
                using (RepairDb db = new RepairDb())
                {
                    if (string.IsNullOrEmpty(model.AreaId))
                    {
                        Int32 seq = db.Query<Int32>("SELECT MAX(sortno) FROM [888_KsSouth].[dbo].[area] ").FirstOrDefault();
                        model.SortNo = seq + 1;
                        sql = string.Format(@" INSERT INTO  [888_KsSouth].[dbo].[area] (SystemCategory,Buliding,Location,SortNo,UpdateUser,UpdateTime)VALUES(@SystemCategory,@Building,@Location,@SortNo,@UpdateUser,@UpdateTime)");
                    }
                    else
                    {
                        sql = string.Format(@" update [888_KsSouth].[dbo].[area] set SystemCategory = @SystemCategory,Buliding = @Building,Location=@Location,UpdateUser=@UpdateUser,UpdateTime = @UpdateTime where area_id = @AreaId ");
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
                    string Msg = null;
                    var modelList = new List<AreaModel>();
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new AreaModel();
                        deleteModel.AreaId = model.AreaId;
                        string checkSql = "SELECT * FROM  [888_KsSouth].[dbo].match WHERE match_type = 'AreaMatch' AND area_id = @AreaId ";
                        var list = db.Query<AreaModel>(checkSql, new { AreaId = model.AreaId });
                        if (list.Count() != 0)
                        {
                            Msg += model.Building+"  ";
                        }
                    }
                    if (!string.IsNullOrEmpty(Msg))
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = Msg + "存在对应的辖区配对，无法删除" }, JsonRequestBehavior.AllowGet);

                    }
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new AreaModel();
                        deleteModel.AreaId = model.AreaId;
                        deleteModel.SystemCategory = model.SystemCategory;
                        deleteModel.Building = model.Building;
                        string sql = string.Format(@" DELETE FROM [888_KsSouth].[dbo].[area]  WHERE area_id = @AreaId ");

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

        //AreaMatch: 辖区配对维护
        public ActionResult AreaMatchIndex()
        {
            return View();
        }

        public ActionResult getAreaMatchList(AreaMatchModel model)
        {
            var list = new List<AreaMatchModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT a.match_type MatchType,
                                                        a.match_id MatchId,
                                                        a.area_id AreaId,
	                                                    b.Buliding Building,
	                                                    b.Location,
                                                        a.charge_emp EmpNo,
	                                                    c.FullName,
                                                        a.sort,
                                                        a.UpdateUser,
                                                        a.UpdateTime,a.SystemCategory
                                                 FROM [888_KsSouth].[dbo].[match] a
                                                     LEFT JOIN [888_KsSouth].[dbo].[area] b
                                                         ON a.area_id = b.area_id
		                                                  LEFT JOIN [888_KsSouth].[dbo].[charge] c
                                                         ON a.charge_emp = c.EmpNo
                                                          and a.SystemCategory = c.SystemCategory
                                                 WHERE 1 = 1 and match_type= 'AreaMatch' ");
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and a.SystemCategory = @SystemCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.Building))
                    {
                        sql += " and b.Buliding = @Building ";
                    }
                    if (!string.IsNullOrEmpty(model.EmpNo))
                    {
                        sql += " and c.EmpNo = @EmpNo ";
                    }
                    var group = Session["GroupName"].ToString();
                    if (!string.IsNullOrEmpty(group))
                    {
                        switch (group)
                        {
                            case "资讯":
                                sql += " and a.SystemCategory = 'IT(资讯类)' ";
                                break;
                            case "后勤":
                                sql += " and a.SystemCategory = 'Logistics(总务后勤类)'";
                                break;
                            default:
                                break;
                        }
                    }
                    sql += " ORDER BY a.sortno asc,a.match_id ASC,a.area_id ASC,a.sort ASC";

                    list = db.Query<AreaMatchModel>(sql, model).ToList();
                }
                return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult editAreaMatch(AreaMatchModel model)
        {
            try
            {
                string sql = "";
                using (RepairDb db = new RepairDb())
                {
                    model.MatchType = "AreaMatch";
                    model.UpdateUser = Session["fullname"].ToString();
                    model.UpdateTime = DateTime.Now;
                    if (string.IsNullOrEmpty(model.MatchId))
                    {
                        string checkSql = @"select * from [888_KsSouth].[dbo].[match] where area_id = @AreaId and charge_emp = @EmpNo and match_type = 'AreaMatch' and SystemCategory = @SystemCategory ";
                        var list = db.Query<AreaMatchModel>(checkSql, model).ToList();
                        if (list.Count() != 0)
                        {
                            return Json(new FlagTips { IsSuccess = false, Msg = "该辖区已维护负责人，请勿重复维护" }, JsonRequestBehavior.AllowGet);
                        }
                        Int32 seq = db.Query<Int32>("SELECT ISNULL(MAX(sortno),'2000') FROM [888_KsSouth].[dbo].[match] WHERE match_type = 'AreaMatch' and SystemCategory = @SystemCategory ", new { model.SystemCategory }).FirstOrDefault();
                        model.SortNo = seq + 1;
                        sql = string.Format(@" INSERT INTO  [888_KsSouth].[dbo].[match] (SystemCategory,match_type,area_id,charge_emp,Sort,SortNo,UpdateUser,UpdateTime)VALUES(@SystemCategory,@MatchType,@AreaId,@EmpNo,@Sort,@SortNo,@UpdateUser,@UpdateTime)");
                    }
                    else
                    {
                        sql = string.Format(@" update [888_KsSouth].[dbo].[match] set SystemCategory = @SystemCategory,area_id = @AreaId,charge_emp=@EmpNo,Sort=@Sort,UpdateUser=@UpdateUser,UpdateTime=@UpdateTime where area_id = @AreaId  and match_type = 'AreaMatch'");
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

        public ActionResult deleteAreaMatch(List<AreaMatchModel> deleteList)
        {
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    var modelList = new List<AreaMatchModel>();
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new AreaMatchModel();
                        deleteModel.MatchId = model.MatchId;
                        deleteModel.SystemCategory = model.SystemCategory;
                        string sql = string.Format(@" DELETE FROM [888_KsSouth].[dbo].[match]  WHERE match_id = @MatchId and match_type = 'AreaMatch' and SystemCategory = @SystemCategory  ");

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