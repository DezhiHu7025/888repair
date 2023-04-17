using _888repair.Db;
using _888repair.Models.State;
using _888repair.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace _888repair.Controllers
{
    public class StateController : Controller
    {
        // GET: State
        public ActionResult StateIndex()
        {
            return View();
        }

        /// <summary>
        /// 状态信息
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult getStateList(StateModel model)

        {
            var list = new List<StateModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT SystemCategory,state_id ID,StatusValue,StatusText,UpdateUser,UpdateTime FROM [888_KsNorth].[dbo].[state]
                                                where 1=1");
                    if (!string.IsNullOrEmpty(model.StatusText))
                    {
                        sql += " and StatusText like '%" + model.StatusText + "%'";
                    }
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and SystemCategory =@SystemCategory ";
                    }
                    sql += " ORDER BY ID asc";

                    list = db.Query<StateModel>(sql, model).ToList();
                }
                return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public ActionResult editStatus(StateModel model)
        {
            try
            {
                string sql = "";
                using (RepairDb db = new RepairDb())
                {
                    model.UpdateUser = "dezhi_hu";
                    model.UpdateTime = DateTime.Now;
                    if (string.IsNullOrEmpty(model.ID))
                    {
                        sql = string.Format(@" INSERT INTO  [888_KsNorth].[dbo].[state] (SystemCategory,StatusValue,StatusText,UpdateUser,UpdateTime)VALUES(@SystemCategory,@StatusValue,@StatusText,@UpdateUser,@UpdateTime)");
                    }
                    else
                    {
                        sql = string.Format(@" update [888_KsNorth].[dbo].[state] set StatusValue = @StatusValue,StatusText=@StatusText,UpdateUser=@UpdateUser,UpdateTime=@UpdateTime where state_id = @ID");
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

        public ActionResult DeleteState(List<StateModel> deleteList)
        {
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    var modelList = new List<StateModel>();
                    foreach (var model in deleteList)
                    {
                        var deleteModel = new StateModel();
                        deleteModel.ID = model.ID;
                        deleteModel.StatusValue = model.StatusValue;

                        string sql = string.Format(@" DELETE FROM [888_KsNorth].[dbo].[state]  WHERE state_id = @ID AND StatusValue = @StatusValue");

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