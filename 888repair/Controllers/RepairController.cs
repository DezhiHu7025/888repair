using _888repair.Db;
using _888repair.Models;
using _888repair.Models.Repair;
using _888repair.Service;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace _888repair.Controllers
{
    public class RepairController : Controller
    {
        // GET: Repair 填写问题
        public ActionResult RepairIndex()
        {
            return View();
        }

        /// <summary>
        /// 总务后勤类报修
        /// </summary>
        /// <returns></returns>
        public ActionResult LogRepairVw()
        {
            return View();
        }

        /// <summary>
        /// IT资讯类报修
        /// </summary>
        /// <returns></returns>
        public ActionResult ITRepairVw()
        {
            return View();
        }

        public ActionResult ITRepairSave(RepairRecordModel model)
        {
            try
            {
                string sql = "";
                model.RepairId = DateTime.Now.ToShortTimeString();
                model.RepairId = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                model.ResponseEmpno = "H22080031";
                model.ResponseEmpname = "dezhi_hu";
                model.CreatTime = DateTime.Now;
                model.Status = "";
                if (!string.IsNullOrEmpty(model.PhotoPath))
                {
                    model.PhotoPath = model.PhotoPath.Substring(0, model.PhotoPath.Length - 1);
                    string[] photos = model.PhotoPath.Split(';');
                    string paths = null;
                    foreach (var photo in photos)
                    {
                        paths += "/UploadFile/" + photo + ";";
                    }
                    model.PhotoPath = paths;
                }
                using (RepairDb db = new RepairDb())
                {
                    var chargeModel = new DirectorModel();
                    //资讯类找顺位为1的辖区负责人  后勤类找顺位为1的业管负责人
                    //抓取负责人工号及姓名
                    if (model.SystemCategory == "IT(资讯类)")
                    {
                        string findChargeSql = @"SELECT a.charge_emp EmpNo,
                                                        b.FullName
                                                FROM[888_KsNorth].[dbo].[match] a
                                                LEFT JOIN[888_KsNorth].[dbo].[charge] b
                                                ON a.charge_emp = b.EmpNo
                                                WHERE a.match_type = 'AreaMatch'
                                                      AND a.SystemCategory = 'IT(资讯类)'
                                                      AND a.area_id = @AreaId
                                                      AND a.sort = '1' ";
                        chargeModel = db.Query<DirectorModel>(findChargeSql, new { AreaId = model.AreaId }).FirstOrDefault();
                        model.ChargeEmpno = chargeModel.EmpNo;
                        model.ChargeEmpname = chargeModel.FullName;
                    }
                    else if (model.SystemCategory == "Logistics(总务后勤类)")
                    {
                        string findChargeSql = @"SELECT a.charge_emp EmpNo,
                                                        b.FullName
                                                FROM[888_KsNorth].[dbo].[match] a
                                                LEFT JOIN[888_KsNorth].[dbo].[charge] b
                                                ON a.charge_emp = b.EmpNo
                                                WHERE a.match_type = 'KindMatch'
                                                      AND a.SystemCategory = 'Logistics(总务后勤类)'
                                                      AND a.area_id = @KindId
                                                      AND a.sort = '1' ";
                        chargeModel = db.Query<DirectorModel>(findChargeSql, new { KindId = model.KindId }).FirstOrDefault();
                    }

                    if (chargeModel == null)
                    {
                        if (!string.IsNullOrEmpty(model.PhotoPath))
                        {
                            FileUploadService upfile = new FileUploadService();
                            string result = upfile.ITRepairPicDelete(model.PhotoPath);
                            if (result != "OK")
                            {
                                return Json(new FlagTips { IsSuccess = false, Msg = "图片处理发生异常，请刷新页面重新提交表单" });
                            }
                        }
                        return Json(new FlagTips { IsSuccess = false, Msg = "负责人抓取异常，请联系资讯" });

                    }
                    sql = @"insert into [888_KsNorth].[dbo].[record] (repair_id,area_id,kind_id,SystemCategory,Building,Loaction,
                                               charge_empno,charge_empname,Category,ResponseContent,ReplyContent,Status,CreatTime,
                                               PhotoPath,RoomNum,repairTime,Telephone,DamageReason,DamageClass,DamageName,ResponseEmpno,
                                               ResponseEmpname,FinishTime) 
                                       values (@RepairId,@AreaId,@KindId,@SystemCategory,@Building,@Loaction,
                                               @ChargeEmpno,@ChargeEmpname,@Category,@ResponseContent,@ReplyContent,@Status,@CreatTime,
                                               @PhotoPath,@RoomNum,@RepairTime,@Telephone,@DamageReason,@DamageClass,@DamageName,@ResponseEmpno,
                                               @ResponseEmpname,@FinishTime)";
                    Dictionary<string, object> trans = new Dictionary<string, object>();
                    trans.Add(sql, model);
                    db.DoExtremeSpeedTransaction(trans);
                }
            }
            catch (Exception ex)
            {
                if (!string.IsNullOrEmpty(model.PhotoPath))
                {
                    FileUploadService upfile = new FileUploadService();
                    string result = upfile.ITRepairPicDelete(model.PhotoPath);
                    if (result != "OK")
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "图片处理发生异常，请刷新页面重新提交表单" });
                    }
                }
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
            return Json(new FlagTips { IsSuccess = true });
        }

        public ActionResult ITRepairPicUpload()
        {
            HttpPostedFileBase httpPostedFileBase = Request.Files["file"];
            ControllerContext.HttpContext.Request.ContentEncoding = Encoding.GetEncoding("UTF-8");
            ControllerContext.HttpContext.Response.Charset = "UTF-8";
            RepairRecordModel model = new RepairRecordModel();
            string Stu_Empno = "H22080031";
            if (httpPostedFileBase != null)
            {
                FileUploadService upfile = new FileUploadService();
                var uploadModel = upfile.FileLoad(httpPostedFileBase, Stu_Empno);
                if (!uploadModel.IsSuccess)
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = "照片上传失败 Photo upload failed" });
                }

            }
            return Json(new FlagTips { IsSuccess = true });
        }


    }
}