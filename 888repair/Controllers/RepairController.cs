using _888repair.Db;
using _888repair.Models;
using _888repair.Models.Repair;
using _888repair.Service;
using Aspose.Cells;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
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

        #region IT资讯类报修
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

                    model.ChargeEmpno = chargeModel.EmpNo;
                    model.ChargeEmpname = chargeModel.FullName;

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
                try
                {
                    string fileName = Path.GetFileName(httpPostedFileBase.FileName);//原始文件名称
                    if (!Directory.Exists(Server.MapPath("~/UploadFile")))
                    {
                        Directory.CreateDirectory(Server.MapPath("~/UploadFile"));
                    }
                    string prefix = DateTime.Now.ToString("yyyyMMdd_") + Stu_Empno + "_";
                    fileName = prefix + fileName;
                    var path = Path.Combine(Server.MapPath("~/UploadFile"), fileName);
                    httpPostedFileBase.SaveAs(path);
                }
                catch (Exception ex)
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
                }

            }
            return Json(new FlagTips { IsSuccess = true });
        }
        #endregion

        #region 个人问题列表
        /// <summary>
        /// 个人问题列表VW
        /// </summary>
        /// <returns></returns>
        public ActionResult RersonalProblemIndex()
        {
            return View();
        }

        /// <summary>
        /// 个人问题列表
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult getRersonalProblemList(RepairRecordModel model)
        {
            model.ResponseEmpno = "H22080031";
            var list = new List<RepairRecordModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT a.repair_id RepairId,a.area_id AreaId,a.kind_id KindId,a.SystemCategory,a.Building,a.Loaction,
                                               a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname,a.Category,a.ResponseContent,a.ReplyContent,b.StatusText Status,a.CreatTime,
                                               a.PhotoPath,RoomNum,a.repairTime RepairTime,a.Telephone,a.DamageReason,a.DamageClass,a.DamageName,a.ResponseEmpno,
                                               a.ResponseEmpname,a.FinishTime from [888_KsNorth].[dbo].[record] a 
                                               left join [888_KsNorth].[dbo].[state] b on a.status = b.StatusValue and a.SystemCategory = b.SystemCategory where 1= 1");
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and a.SystemCategory =@SystemCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.ResponseEmpno))
                    {
                        sql += " and a.ResponseEmpno =@ResponseEmpno ";
                    }
                    if (!string.IsNullOrEmpty(model.Status))
                    {
                        sql += " and a.Status =@Status ";
                    }
                    if (model.startDate != null)
                    {
                        sql += " and a.CreatTime >= @startDate ";
                    }
                    if (model.endDate != null)
                    {
                        model.endDate = Convert.ToDateTime(model.endDate).AddDays(1);
                        sql += " and a.CreatTime <=@endDate ";
                    }
                    sql += " ORDER BY repair_id desc";

                    list = db.Query<RepairRecordModel>(sql, model).ToList();
                }
                return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public List<RepairRecordModel> queryRersonalProblemList(RepairRecordModel model)
        {
            model.ResponseEmpno = "H22080031";
            var list = new List<RepairRecordModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT a.repair_id RepairId,a.area_id AreaId,a.kind_id KindId,a.SystemCategory,a.Building,a.Loaction,
                                               a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname,a.Category,a.ResponseContent,a.ReplyContent,b.StatusText Status,a.CreatTime,
                                               a.PhotoPath,RoomNum,a.repairTime RepairTime,a.Telephone,a.DamageReason,a.DamageClass,a.DamageName,a.ResponseEmpno,
                                               a.ResponseEmpname,a.FinishTime from [888_KsNorth].[dbo].[record] a 
                                               left join [888_KsNorth].[dbo].[state] b on a.status = b.StatusValue and a.SystemCategory = b.SystemCategory where 1= 1");
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and a.SystemCategory =@SystemCategory ";
                    }
                    if (!string.IsNullOrEmpty(model.ResponseEmpno))
                    {
                        sql += " and a.ResponseEmpno =@ResponseEmpno ";
                    }
                    if (!string.IsNullOrEmpty(model.Status))
                    {
                        sql += " and a.Status =@Status ";
                    }
                    if (model.startDate != null)
                    {
                        sql += " and a.CreatTime >= @startDate ";
                    }
                    if (model.endDate != null)
                    {
                        model.endDate = Convert.ToDateTime(model.endDate).AddDays(1);
                        sql += " and a.CreatTime <=@endDate ";
                    }
                    sql += " ORDER BY repair_id desc";

                    list = db.Query<RepairRecordModel>(sql, model).ToList();
                }
            }
            catch (Exception ex)
            {

            }
            return list;
        }

        /// <summary>
        /// 个人问题列表下载
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult RersonalProblemExport(RepairRecordModel model)
        {
            try
            {
                var dt = queryRersonalProblemList(model);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string templatePath = string.Format("~\\Excel\\RersonalProblem.xlsx");
                FileStream fs = new FileStream(System.Web.HttpContext.Current.Server.MapPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                Workbook wb = new Workbook(fs);
                Worksheet sheet = wb.Worksheets[0];
                sheet.Name = "个人问题列表";
                Cells cells = sheet.Cells;
                int columnCount = cells.MaxColumn;  //获取表页的最大列数
                int rowCount = cells.MaxRow;        //获取表页的最大行数

                for (int col = 0; col < columnCount; col++)
                {
                    sheet.AutoFitColumn(col, 0, rowCount);
                }
                for (int col = 0; col < columnCount; col++)
                {
                    cells.SetColumnWidthPixel(col, cells.GetColumnWidthPixel(col) + 30);
                }

                for (int i = 0; i < dt.Count; i++)//遍历DataTable行
                {
                    sheet.Cells[i + 1, 0].PutValue(dt[i].SystemCategory);
                    sheet.Cells[i + 1, 1].PutValue(dt[i].Building);
                    sheet.Cells[i + 1, 2].PutValue(dt[i].Loaction);
                    sheet.Cells[i + 1, 3].PutValue(dt[i].Category);
                    sheet.Cells[i + 1, 4].PutValue(dt[i].ChargeEmpname);
                    sheet.Cells[i + 1, 5].PutValue(dt[i].ResponseContent);
                    sheet.Cells[i + 1, 6].PutValue(dt[i].Status);
                    sheet.Cells[i + 1, 7].PutValue(dt[i].CreatTime.ToString() == "0001/01/01" ? "" : dt[i].CreatTime.ToString());
                    sheet.Cells[i + 1, 8].PutValue(dt[i].FinishTime.ToString() == "0001/01/01" ? "" : dt[i].FinishTime.ToString());

                    //  sheet.Cells[i + 1, 24].PutValue(Convert.ToDateTime(dt[i].CreateTime).ToString("yyyy/MM/dd") == "0001/01/01" ? "" : Convert.ToDateTime(dt[i].CreateTime).ToString("yyyy-MM-dd HH:mm:ss"));
                }
                MemoryStream bookStream = new MemoryStream();//创建文件流
                wb.Save(bookStream, new OoxmlSaveOptions(SaveFormat.Xlsx)); //文件写入流（向流中写入字节序列）
                bookStream.Seek(0, SeekOrigin.Begin);//输出之前调用Seek，把0位置指定为开始位置
                return File(bookStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("RersonalProblem_{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff")));//最后以文件形式返回
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
        }
        #endregion
    }
}