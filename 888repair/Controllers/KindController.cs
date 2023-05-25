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
    [App_Start.AuthFilter]
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
                    model.UpdateUser = Session["fullname"].ToString();
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
                    string Msg = null;
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new KindModel();
                        deleteModel.KindID = model.KindID;
                        string checkSql = "SELECT * FROM  [888_KsNorth].[dbo].match WHERE match_type = 'KindMatch' AND area_id = @KindID ";
                        var list = db.Query<KindModel>(checkSql, new { KindID = model.KindID });
                        if (list.Count() != 0)
                        {
                            Msg += model.KindCategory + "  ";
                        }
                    }
                    if (!string.IsNullOrEmpty(Msg))
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = Msg + "存在对应的业管配对，无法删除" }, JsonRequestBehavior.AllowGet);

                    }
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new KindModel();
                        deleteModel.KindID = model.KindID;
                        string sql = string.Format(@" DELETE FROM [888_KsNorth].[dbo].[kind]  WHERE kind_id = @KindID ");

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

        //KindMatch: 业管配对维护
        public ActionResult KindMatchIndex()
        {
            return View();
        }

        public ActionResult getKindMatchList(KindMatchModel model)
        {
            var list = new List<KindMatchModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT a.match_type MatchType,
                                                        a.match_id MatchId,
                                                        a.area_id AreaId,
	                                                    b.KindCategory,
	                                                    b.Remark,
                                                        a.charge_emp EmpNo,
	                                                    c.FullName,
                                                        a.sort,
                                                        a.UpdateUser,
                                                        a.UpdateTime,a.SystemCategory
                                                 FROM [888_KsNorth].[dbo].[match] a
                                                     LEFT JOIN [888_KsNorth].[dbo].[kind] b
                                                         ON a.area_id = b.kind_id
		                                                  LEFT JOIN [888_KsNorth].[dbo].[charge] c
                                                         ON a.charge_emp = c.EmpNo
                                                 WHERE 1 = 1 and match_type= 'KindMatch' ");
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and a.SystemCategory = @SystemCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.KindCategory))
                    {
                        sql += " and a.area_id = @KindCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.EmpNo))
                    {
                        sql += " and c.EmpNo = @EmpNo ";
                    }
                    sql += " ORDER BY a.sortno asc,a.match_id ASC,a.area_id ASC,a.sort ASC";

                    list = db.Query<KindMatchModel>(sql, model).ToList();
                }
                return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult editKindMatch(KindMatchModel model)
        {
            try
            {
                string sql = "";
                using (RepairDb db = new RepairDb())
                {
                    model.MatchType = "KindMatch";
                    model.UpdateUser = Session["fullname"].ToString();
                    model.UpdateTime = DateTime.Now;
                    if (string.IsNullOrEmpty(model.MatchId))
                    {
                        string checkSql = @"select * from [888_KsNorth].[dbo].[match] where area_id = @AreaId and charge_emp = @EmpNo and  match_type = 'KindMatch' and SystemCategory = @SystemCategory";
                        var list = db.Query<KindMatchModel>(checkSql, model).ToList();
                        if (list.Count() != 0)
                        {
                            return Json(new FlagTips { IsSuccess = false, Msg = "该辖区已维护负责人，请勿重复维护" }, JsonRequestBehavior.AllowGet);
                        }
                        Int32 seq = db.Query<Int32>("SELECT MAX(sortno) FROM [888_KsNorth].[dbo].[match] WHERE match_type = 'KindMatch' and  SystemCategory = @SystemCategory ",new { model.SystemCategory }).FirstOrDefault();
                        model.SortNo = seq + 1;
                        sql = string.Format(@" INSERT INTO  [888_KsNorth].[dbo].[match] (SystemCategory,match_type,area_id,charge_emp,Sort,SortNo,UpdateUser,UpdateTime)VALUES(@SystemCategory,@MatchType,@AreaId,@EmpNo,@Sort,@SortNo,@UpdateUser,@UpdateTime)");
                    }
                    else
                    {
                        sql = string.Format(@" update [888_KsNorth].[dbo].[match] set SystemCategory=@SystemCategory,area_id = @AreaId,charge_emp=@EmpNo,Sort=@Sort,UpdateUser=@UpdateUser,UpdateTime=@UpdateTime where area_id = @AreaId and match_type ='KindMatch' ");
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

        public ActionResult deleteKindMatch(List<KindMatchModel> deleteList)
        {
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    var modelList = new List<KindMatchModel>();
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new KindMatchModel();
                        deleteModel.SystemCategory = model.SystemCategory;
                        deleteModel.MatchId = model.MatchId;
                        string sql = string.Format(@" DELETE FROM [888_KsNorth].[dbo].[match]  WHERE SystemCategory= @SystemCategory and match_id = @MatchId  and match_type = 'KindMatch' ");

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