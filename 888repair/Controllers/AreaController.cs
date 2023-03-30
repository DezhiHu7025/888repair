﻿using _888repair.Db;
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
                model.UpdateUser = "dezhi_hu";
                model.UpdateTime = DateTime.Now;
                using (RepairDb db = new RepairDb())
                {
                    if (string.IsNullOrEmpty(model.AreaId))
                    {
                        Int32 seq = db.Query<Int32>("SELECT MAX(sortno) FROM [888_KsNorth].[dbo].[area] ").FirstOrDefault();
                        model.SortNo = seq + 1;
                        sql = string.Format(@" INSERT INTO  [888_KsNorth].[dbo].[area] (SystemCategory,Buliding,Location,SortNo,UpdateUser,UpdateTime)VALUES(@SystemCategory,@Buliding,@Location,@SortNo,@UpdateUser,@UpdateTime)");
                    }
                    else
                    {
                        sql = string.Format(@" update [888_KsNorth].[dbo].[area] set SystemCategory = @SystemCategory,Buliding = @Buliding,Location=@Location,UpdateUser=@UpdateUser,UpdateTime = @UpdateTime where area_id = @AreaId ");
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

        //AreaMatchI: 辖区配对维护
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
	                                                    b.Buliding,
	                                                    b.Location,
                                                        a.charge_emp EmpNo,
	                                                    c.FullName,
                                                        a.sort,
                                                        a.UpdateUser,
                                                        a.UpdateTime
                                                 FROM [888_KsNorth].[dbo].[match] a
                                                     LEFT JOIN [888_KsNorth].[dbo].[area] b
                                                         ON a.area_id = b.area_id
		                                                  LEFT JOIN [888_KsNorth].[dbo].[charge] c
                                                         ON a.charge_emp = c.EmpNo
                                                 WHERE 1 = 1 and match_type= 'AreaMatch' ");
                    if (!string.IsNullOrEmpty(model.Buliding))
                    {
                        sql += " and a.area_id = @Buliding ";
                    }
                    if (!string.IsNullOrEmpty(model.EmpNo))
                    {
                        sql += " and c.EmpNo = @EmpNo ";
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
                    model.UpdateUser = "dezhi_hu";
                    model.UpdateTime = DateTime.Now;
                    if (string.IsNullOrEmpty(model.MatchId))
                    {
                        string checkSql = @"select * from [888_KsNorth].[dbo].[match] where area_id = @AreaId and charge_emp = @EmpNo ";
                        var list = db.Query<AreaMatchModel>(checkSql, model).ToList();
                        if (list.Count() != 0)
                        {
                            return Json(new FlagTips { IsSuccess = false, Msg = "该辖区已维护负责人，请勿重复维护" }, JsonRequestBehavior.AllowGet);
                        }
                        Int32 seq = db.Query<Int32>("SELECT MAX(sortno) FROM [888_KsNorth].[dbo].[match] WHERE match_type = 'AreaMatch' ").FirstOrDefault();
                        model.SortNo = seq + 1;
                        sql = string.Format(@" INSERT INTO  [888_KsNorth].[dbo].[match] (match_type,area_id,charge_emp,Sort,SortNo,UpdateUser,UpdateTime)VALUES(@MatchType,@AreaId,@EmpNo,@Sort,@SortNo,@UpdateUser,@UpdateTime)");
                    }
                    else
                    {
                        sql = string.Format(@" update [888_KsNorth].[dbo].[match] set area_id = @AreaId,charge_emp=@EmpNo,Sort=@Sort,UpdateUser=@UpdateUser,UpdateTime=@UpdateTime where area_id = @AreaId ");
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
                        deleteModel.AreaId = model.AreaId;
                        deleteModel.MatchId = model.MatchId;
                        deleteModel.MatchType = model.MatchType;
                        string sql = string.Format(@" DELETE FROM [888_KsNorth].[dbo].[match]  WHERE match_id = @MatchId ");

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