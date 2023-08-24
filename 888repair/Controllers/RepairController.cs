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
    [App_Start.AuthFilter]
    public class RepairController : Controller
    {
        /// <summary>
        /// 移动端
        /// </summary>
        /// <returns></returns>
        public ActionResult RepairMobileIndex()
        {
            return View();
        }

        public ActionResult LogRepairMobileVw()
        {
            return View();
        }

        public ActionResult ITRepairMobileVw()
        {
            return View();
        }

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
                model.RepairId = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                model.ResponseEmpno = Session["EmpNo"].ToString();
                model.ResponseEmpname = Session["fullname"].ToString();
                model.CreatTime = DateTime.Now;
                model.Status = "pending";
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

                    StepRecordModel stepModel = new StepRecordModel();
                    stepModel.GUID = Guid.NewGuid().ToString();
                    stepModel.RepairId = model.RepairId;
                    stepModel.STATUS = model.Status;
                    stepModel.STEP = "送出报修单";
                    stepModel.ChargeEmpno = model.ChargeEmpno;
                    stepModel.ChargeEmpname = model.ChargeEmpname;
                    stepModel.UpdateEmpNo = model.ResponseEmpno;
                    stepModel.UpdateEmpName = model.ResponseEmpname;
                    stepModel.UpdateTime = model.CreatTime;
                    string stepSQL = @"insert into [888_KsNorth].[dbo].[steprecord] ([guid],[repair_id],[status],[step],[charge_empno],[charge_empname],[UpdateEmpNo],[UpdateEmpName],[UpdateTime]) 
                                            values (@GUID,@RepairId,@STATUS,@STEP,@ChargeEmpno,@ChargeEmpname,@UpdateEmpNo,@UpdateEmpName,@UpdateTime)";

                    SendMailService sendMailService = new SendMailService();
                    bool mailFlag = sendMailService.newMail(model.AreaId, model.KindId, model.RepairId);
                    if (mailFlag)
                    {
                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(sql, model);
                        trans.Add(stepSQL, stepModel);
                        db.DoExtremeSpeedTransaction(trans);
                    }
                    else
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "邮件通知异常，送出报修单失败" });
                    }

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

        public ActionResult ITRepairPicUpload(string datestring)
        {
            HttpPostedFileBase httpPostedFileBase = Request.Files["file"];
            ControllerContext.HttpContext.Request.ContentEncoding = Encoding.GetEncoding("UTF-8");
            ControllerContext.HttpContext.Response.Charset = "UTF-8";
            RepairRecordModel model = new RepairRecordModel();
            string Stu_Empno = Session["EmpNo"].ToString();
            if (httpPostedFileBase != null)
            {
                try
                {
                    string fileName = Path.GetFileName(httpPostedFileBase.FileName);//原始文件名称
                    if (!Directory.Exists(Server.MapPath("~/UploadFile")))
                    {
                        Directory.CreateDirectory(Server.MapPath("~/UploadFile"));
                    }
                    // string prefix = DateTime.Now.ToString("yyyyMMddHH_") + Stu_Empno + "_";
                    string prefix = datestring + "_" + Stu_Empno + "_";
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

        /// <summary>
        /// 完修图片上传
        /// </summary>
        /// <returns></returns>
        public ActionResult ReplyPicUpload(string datestring)
        {
            HttpPostedFileBase httpPostedFileBase = Request.Files["file"];
            ControllerContext.HttpContext.Request.ContentEncoding = Encoding.GetEncoding("UTF-8");
            ControllerContext.HttpContext.Response.Charset = "UTF-8";
            RepairRecordModel model = new RepairRecordModel();
            string Stu_Empno = Session["EmpNo"].ToString();
            if (httpPostedFileBase != null)
            {
                try
                {
                    string fileName = Path.GetFileName(httpPostedFileBase.FileName);//原始文件名称
                    if (!Directory.Exists(Server.MapPath("~/ReplyPhoto")))
                    {
                        Directory.CreateDirectory(Server.MapPath("~/ReplyPhoto"));
                    }
                    //string prefix = DateTime.Now.ToString("yyyyMMddHH_") + Stu_Empno + "_";
                    string prefix = datestring + "_" + Stu_Empno + "_";
                    fileName = prefix + fileName;
                    var path = Path.Combine(Server.MapPath("~/ReplyPhoto"), fileName);
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
        /// 个人问题列表VW移动端
        /// </summary>
        /// <returns></returns>
        public ActionResult RPMobileIndex()
        {
            return View();
        }

        /// <summary>
        /// 报修单详情移动端
        /// </summary>
        /// <returns></returns>
        public ActionResult RepairDetailMobileVw()
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
            model.ResponseEmpno = Session["EmpNo"].ToString();
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
                    //执行状态
                    if (!string.IsNullOrEmpty(model.Status))
                    {
                        model.queryStatus = model.Status.Split(',');
                        sql += " and a.status in @queryStatus ";
                    }
                    //填表时间
                    if (model.queryYear != null && model.queryMonths != null)
                    {
                        string[] monthsString = model.queryMonths.Split(',');
                        List<string> queryDate = new List<string>();
                        foreach (var item in monthsString)
                        {
                            queryDate.Add(model.queryYear + item);
                        }
                        model.queryDates = queryDate.ToArray();
                        sql += " and CONCAT(YEAR(a.CreatTime),MONTH(a.CreatTime)) in  @queryDates";
                    }
                    if (model.queryYear != null && model.queryMonths == null)
                    {
                        sql += " and YEAR(a.CreatTime) =@queryYear ";
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
            model.ResponseEmpno = Session["EmpNo"].ToString();
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
                    //执行状态
                    if (!string.IsNullOrEmpty(model.Status))
                    {
                        model.queryStatus = model.Status.Split(',');
                        sql += " and a.status in @queryStatus ";
                    }
                    //填表时间
                    if (model.queryYear != null && model.queryMonths != null)
                    {
                        string[] monthsString = model.queryMonths.Split(',');
                        List<string> queryDate = new List<string>();
                        foreach (var item in monthsString)
                        {
                            queryDate.Add(model.queryYear + item);
                        }
                        model.queryDates = queryDate.ToArray();
                        sql += " and CONCAT(YEAR(a.CreatTime),MONTH(a.CreatTime)) in  @queryDates";
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
                string templatePath = string.Format("~\\Excel\\RersonalProblemList.xlsx");
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
                    sheet.Cells[i + 1, 0].PutValue(dt[i].RepairId);//请修编号
                    sheet.Cells[i + 1, 1].PutValue(dt[i].SystemCategory);//系统类别
                    sheet.Cells[i + 1, 2].PutValue(dt[i].Building);//大楼别
                    sheet.Cells[i + 1, 3].PutValue(dt[i].Loaction);//位置
                    sheet.Cells[i + 1, 4].PutValue(dt[i].RoomNum);//空间编号
                    sheet.Cells[i + 1, 5].PutValue(dt[i].ChargeEmpname);//负责人姓名
                    sheet.Cells[i + 1, 6].PutValue(dt[i].ResponseEmpname);//反应人姓名
                    sheet.Cells[i + 1, 7].PutValue(dt[i].Telephone);//分机号码
                    sheet.Cells[i + 1, 8].PutValue(dt[i].RepairTime);//维护时间
                    sheet.Cells[i + 1, 9].PutValue(dt[i].CreatTime.ToString() == "0001/01/01" ? "" : dt[i].CreatTime.ToString());//填表时间

                    sheet.Cells[i + 1, 10].PutValue(dt[i].ResponseContent);//反应内容
                    sheet.Cells[i + 1, 11].PutValue(dt[i].Status);//执行状态
                    sheet.Cells[i + 1, 12].PutValue(Convert.ToDateTime(dt[i].FinishTime).ToString("yyyy/MM/dd") == "0001/01/01" ? "" : Convert.ToDateTime(dt[i].FinishTime).ToString("yyyy-MM-dd HH:mm:ss"));
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

        #region 报修单 回复问题
        /// <summary>
        /// 报修单详细信息
        /// </summary>
        /// <returns></returns>
        public ActionResult RepairDetailVw()
        {
            return View();
        }

        /// <summary>
        /// 根据请修编号获取报修单详情
        /// </summary>
        /// <param name="RepairId"></param>
        /// <returns></returns>
        public ActionResult getRepairDetail(string RepairId)
        {
            var model = new RepairRecordModel();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT a.repair_id RepairId,a.area_id AreaId,a.kind_id KindId,a.SystemCategory,a.Building,a.Loaction,
                                               a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname,a.Category,a.ResponseContent,a.ReplyContent,a.Status,a.CreatTime,
                                               a.PhotoPath,RoomNum,a.repairTime RepairTime,a.Telephone,a.DamageReason,a.DamageClass,a.DamageName,a.ResponseEmpno,
                                               a.ResponseEmpname,a.FinishTime,a.ReplyPhoto from [888_KsNorth].[dbo].[record] a 
                                               where 1= 1");
                    if (!string.IsNullOrEmpty(RepairId))
                    {
                        sql += " and a.repair_id =@RepairId ";
                    }

                    model = db.Query<RepairRecordModel>(sql, new { RepairId }).FirstOrDefault();
                }
                return Json(model, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        /// <summary>
        /// 回复问题
        /// </summary>
        /// <returns></returns>
        public ActionResult ReplyRepairIndex()
        {
            return View();
        }

        /// <summary>
        /// 回复问题报修单详情
        /// </summary>
        /// <returns></returns>
        public ActionResult ReplyRepairDetailVw()
        {
            return View();
        }

        /// <summary>
        /// 回复问题列表查询
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult getReplyProblemList(RepairRecordModel model)
        {
            model.ResponseEmpno = Session["EmpNo"].ToString();
            model.EmpNo = Session["EmpNo"].ToString();
            var list = new List<RepairRecordModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT DISTINCT a.repair_id RepairId,a.area_id AreaId,a.kind_id KindId,a.SystemCategory,a.Building Building,a.Loaction,
                                               a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname,a.Category,a.ResponseContent,a.ReplyContent,b.StatusText Status,a.CreatTime,
                                               a.PhotoPath,RoomNum,a.repairTime RepairTime,a.Telephone,a.DamageReason,a.DamageClass,a.DamageName,a.ResponseEmpno,
                                               a.ResponseEmpname,a.FinishTime from [888_KsNorth].[dbo].[record] a 
                                               left join [888_KsNorth].[dbo].[state] b on a.status = b.StatusValue and a.SystemCategory = b.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] area  ON a.area_id = area.area_id AND area.match_type = 'AreaMatch' AND a.SystemCategory = area.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] kind  ON a.kind_id = kind.area_id AND kind.match_type = 'KindMatch' AND a.SystemCategory = kind.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch1  ON area.charge_emp = ch1.EmpNo AND a.SystemCategory = ch1.SystemCategory 
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch2  ON area.charge_emp = ch2.EmpNo AND a.SystemCategory = ch2.SystemCategory WHERE 1= 1
											   AND (area.charge_emp = @EmpNo OR kind.charge_emp = @EmpNo OR a.charge_empno = @EmpNo)");
                    //系统类别
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and a.SystemCategory =@SystemCategory ";
                    }
                    //大楼别
                    if (!string.IsNullOrEmpty(model.Building))
                    {
                        sql += " and a.Building =@Building ";
                    }
                    //位置
                    if (!string.IsNullOrEmpty(model.Loaction))
                    {
                        sql += " and a.Loaction =@Loaction ";
                    }
                    //类别
                    if (!string.IsNullOrEmpty(model.Category))
                    {
                        sql += " and a.Category =@Category ";
                    }
                    //反应人姓名
                    if (!string.IsNullOrEmpty(model.ResponseEmpname))
                    {
                        sql += " and a.ResponseEmpname like '%" + model.ResponseEmpname + "%'";
                    }
                    //反应内容
                    if (!string.IsNullOrEmpty(model.ResponseContent))
                    {
                        sql += " and a.ResponseContent like '%" + model.ResponseContent + "%'";
                    }
                    //执行状态
                    if (!string.IsNullOrEmpty(model.Status))
                    {
                        sql += " and a.Status =@Status ";
                    }
                    //负责人
                    if (!string.IsNullOrEmpty(model.ChargeEmpname))
                    {
                        sql += " and a.charge_empname =@ChargeEmpname ";
                    }
                    //填表时间
                    if (model.startDate != null)
                    {
                        sql += " and a.CreatTime >= @startDate ";
                    }
                    if (model.endDate != null)
                    {
                        model.endDate = Convert.ToDateTime(model.endDate).AddDays(1);
                        sql += " and a.CreatTime <=@endDate ";
                    }
                    var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
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

        public List<RepairRecordModel> queryReplyProblemList(RepairRecordModel model)
        {
            model.ResponseEmpno = Session["EmpNo"].ToString();
            model.EmpNo = Session["EmpNo"].ToString();
            var list = new List<RepairRecordModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT DISTINCT a.repair_id RepairId,a.area_id AreaId,a.kind_id KindId,a.SystemCategory,a.Building Building,a.Loaction,
                                               a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname,a.Category,a.ResponseContent,a.ReplyContent,b.StatusText Status,a.CreatTime,
                                               a.PhotoPath,RoomNum,a.repairTime RepairTime,a.Telephone,a.DamageReason,a.DamageClass,a.DamageName,a.ResponseEmpno,
                                               a.ResponseEmpname,a.FinishTime from [888_KsNorth].[dbo].[record] a 
                                               left join [888_KsNorth].[dbo].[state] b on a.status = b.StatusValue and a.SystemCategory = b.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] area  ON a.area_id = area.area_id AND area.match_type = 'AreaMatch' AND a.SystemCategory = area.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] kind  ON a.kind_id = kind.area_id AND kind.match_type = 'KindMatch' AND a.SystemCategory = kind.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch1  ON area.charge_emp = ch1.EmpNo AND a.SystemCategory = ch1.SystemCategory 
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch2  ON area.charge_emp = ch2.EmpNo AND a.SystemCategory = ch2.SystemCategory WHERE 1= 1
											   AND (area.charge_emp = @EmpNo OR kind.charge_emp = @EmpNo OR a.charge_empno = @EmpNo)");
                    //系统类别
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and a.SystemCategory =@SystemCategory ";
                    }
                    //大楼别
                    if (!string.IsNullOrEmpty(model.Building))
                    {
                        sql += " and a.Building =@Building ";
                    }
                    //位置
                    if (!string.IsNullOrEmpty(model.Loaction))
                    {
                        sql += " and a.Loaction =@Loaction ";
                    }
                    //类别
                    if (!string.IsNullOrEmpty(model.Category))
                    {
                        sql += " and a.Category =@Category ";
                    }
                    //反应人姓名
                    if (!string.IsNullOrEmpty(model.ResponseEmpname))
                    {
                        sql += " and a.ResponseEmpname like '%" + model.ResponseEmpname + "%'";
                    }
                    //反应内容
                    if (!string.IsNullOrEmpty(model.ResponseContent))
                    {
                        sql += " and a.ResponseContent like '%" + model.ResponseContent + "%'";
                    }
                    //执行状态
                    if (!string.IsNullOrEmpty(model.Status))
                    {
                        sql += " and a.Status =@Status ";
                    }
                    //负责人
                    if (!string.IsNullOrEmpty(model.ChargeEmpname))
                    {
                        sql += " and a.charge_empname =@ChargeEmpname ";
                    }
                    //填表时间
                    if (model.startDate != null)
                    {
                        sql += " and a.CreatTime >= @startDate ";
                    }
                    if (model.endDate != null)
                    {
                        model.endDate = Convert.ToDateTime(model.endDate).AddDays(1);
                        sql += " and a.CreatTime <=@endDate ";
                    }
                    var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
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
                    sql += " ORDER BY repair_id desc";

                    list = db.Query<RepairRecordModel>(sql, model).ToList();
                }
              //  return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
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
        public ActionResult ReplyProblemListExport(RepairRecordModel model)
        {
            try
            {
                var dt = queryReplyProblemList(model);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string templatePath = string.Format("~\\Excel\\ProblemList.xlsx");
                FileStream fs = new FileStream(System.Web.HttpContext.Current.Server.MapPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                Workbook wb = new Workbook(fs);
                Worksheet sheet = wb.Worksheets[0];
                sheet.Name = "回复问题";
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
                    sheet.Cells[i + 1, 0].PutValue(dt[i].RepairId);
                    sheet.Cells[i + 1, 1].PutValue(dt[i].SystemCategory);
                    sheet.Cells[i + 1, 2].PutValue(dt[i].Building);
                    sheet.Cells[i + 1, 3].PutValue(dt[i].Loaction);
                    sheet.Cells[i + 1, 4].PutValue(dt[i].Category);
                    sheet.Cells[i + 1, 5].PutValue(dt[i].ResponseEmpname);
                    sheet.Cells[i + 1, 6].PutValue(dt[i].ChargeEmpname);
                    sheet.Cells[i + 1, 7].PutValue(dt[i].ResponseContent);
                    sheet.Cells[i + 1, 8].PutValue(dt[i].Status);
                    sheet.Cells[i + 1, 9].PutValue(dt[i].CreatTime.ToString() == "0001/01/01" ? "" : dt[i].CreatTime.ToString());
                    sheet.Cells[i + 1, 10].PutValue(dt[i].FinishTime.ToString() == "0001/01/01" ? "" : dt[i].FinishTime.ToString());
                }
                MemoryStream bookStream = new MemoryStream();//创建文件流
                wb.Save(bookStream, new OoxmlSaveOptions(SaveFormat.Xlsx)); //文件写入流（向流中写入字节序列）
                bookStream.Seek(0, SeekOrigin.Begin);//输出之前调用Seek，把0位置指定为开始位置
                return File(bookStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("ReplyProblemList_{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff")));//最后以文件形式返回
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
        }

        /// <summary>
        /// 回复报修单
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult ReplyProblemSave(RepairRecordModel model)
        {

            model.EmpNo = Session["EmpNo"].ToString();
            model.FullName = Session["fullname"].ToString();
            if (!string.IsNullOrEmpty(model.ReplyPhoto))
            {
                model.ReplyPhoto = model.ReplyPhoto.Substring(0, model.ReplyPhoto.Length - 1);
                string[] photos = model.ReplyPhoto.Split(';');
                string paths = null;
                foreach (var photo in photos)
                {
                    paths += "/ReplyPhoto/" + photo + ";";
                }
                model.ReplyPhoto = paths;
            }

            StepRecordModel stepModel = new StepRecordModel();
            stepModel.GUID = Guid.NewGuid().ToString();
            stepModel.OPINION = model.NewReplyContent;//获取未重新定义前的新回复内容
            stepModel.RepairId = model.RepairId;
            stepModel.STATUS = model.Status;
            stepModel.STEP = "回应报修单";
            stepModel.ChargeEmpno = model.ChargeEmpno;
            stepModel.ChargeEmpname = model.ChargeEmpname;
            stepModel.UpdateEmpNo = model.EmpNo;
            stepModel.UpdateEmpName = model.FullName;
            stepModel.UpdateTime = DateTime.Now;

            model.NewReplyContent = model.NewReplyContent + "\r\n-" + model.FullName + "于" + DateTime.Now + "回应-" + "\r\n\r\n";
            if (model.Status == "complete" || model.Status == "rejectend")
            {
                model.FinishTime = DateTime.Now;
            }
            try
            {
                model.ResponseEmpno = Session["EmpNo"].ToString();
                model.ResponseEmpname = Session["fullname"].ToString();
                model.CreatTime = DateTime.Now;
                using (RepairDb db = new RepairDb())
                {
                    string sql = @" update [888_KsNorth].[dbo].[record]  set ReplyContent = concat(ReplyContent,@NewReplyContent),Status = @Status,FinishTime = @FinishTime,ReplyPhoto=@ReplyPhoto where repair_id = @RepairId ";

                    string stepSQL = @"insert into [888_KsNorth].[dbo].[steprecord] ([guid],[repair_id],[status],[OPINION],[step],[charge_empno],[charge_empname],[UpdateEmpNo],[UpdateEmpName],[UpdateTime]) 
                                            values (@GUID,@RepairId,@STATUS,@OPINION,@STEP,@ChargeEmpno,@ChargeEmpname,@UpdateEmpNo,@UpdateEmpName,@UpdateTime)";
                    SendMailService sendMailService = new SendMailService();
                    bool mailFlag = sendMailService.replyMail(model.RepairId, model.Status);
                    if (mailFlag)
                    {
                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(sql, model);
                        trans.Add(stepSQL, stepModel);
                        db.DoExtremeSpeedTransaction(trans);
                    }
                    else
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "邮件通知异常，回应报修单失败" });
                    }

                }
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
            return Json(new FlagTips { IsSuccess = true });
        }

        /// <summary>
        /// 报修单转派
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult TransferProblem(RepairRecordModel model)
        {

            model.EmpNo = Session["EmpNo"].ToString();
            model.FullName = Session["fullname"].ToString();

            StepRecordModel stepModel = new StepRecordModel();
            stepModel.GUID = Guid.NewGuid().ToString();
            stepModel.OPINION = model.NewReplyContent;//获取未重新定义前的新回复内容
            stepModel.RepairId = model.RepairId;
            //stepModel.STATUS = model.Status;
            stepModel.STEP = "转派报修单";
            stepModel.ChargeEmpno = model.ChargeEmpno;
            stepModel.ChargeEmpname = model.ChargeEmpname;
            stepModel.UpdateEmpNo = model.EmpNo;
            stepModel.UpdateEmpName = model.FullName;
            stepModel.UpdateTime = DateTime.Now;

            model.NewReplyContent = model.NewReplyContent + "\r\n-" + model.FullName + "于" + DateTime.Now + "转派给-" + model.ChargeEmpname + "\r\n\r\n";
            try
            {
                string sql = "";
                using (RepairDb db = new RepairDb())
                {
                    string findStatusSql = @"SELECT Status FROM  [888_KsNorth].[dbo].[record] WHERE repair_id = @RepairId";
                    stepModel.STATUS = db.Query<string>(findStatusSql, new { RepairId = model.RepairId }).FirstOrDefault();
                    sql = @" update [888_KsNorth].[dbo].[record]  set ReplyContent = concat(ReplyContent,@NewReplyContent),charge_empno = @ChargeEmpno,charge_empname = @ChargeEmpname where repair_id = @RepairId ";

                    string stepSQL = @"insert into [888_KsNorth].[dbo].[steprecord] ([guid],[repair_id],[status],[step],[OPINION],[charge_empno],[charge_empname],[UpdateEmpNo],[UpdateEmpName],[UpdateTime]) 
                                            values (@GUID,@RepairId,@STATUS,@STEP,@OPINION,@ChargeEmpno,@ChargeEmpname,@UpdateEmpNo,@UpdateEmpName,@UpdateTime)";

                    SendMailService sendMailService = new SendMailService();
                    bool mailFlag = sendMailService.transferMail(model.RepairId, model.ChargeEmpno);
                    if (mailFlag)
                    {
                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(sql, model);
                        trans.Add(stepSQL, stepModel);
                        db.DoExtremeSpeedTransaction(trans);
                    }
                    else
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "邮件通知异常，转派报修单失败" });
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
            return Json(new FlagTips { IsSuccess = true });
        }

        /// <summary>
        /// 报修单 驳回任务
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult RejectTask(RepairRecordModel model)
        {
            model.EmpNo = Session["EmpNo"].ToString();
            model.FullName = Session["fullname"].ToString();

            StepRecordModel stepModel = new StepRecordModel();
            stepModel.GUID = Guid.NewGuid().ToString();
            stepModel.OPINION = model.NewReplyContent;//获取未重新定义前的新回复内容
            stepModel.RepairId = model.RepairId;
            stepModel.STEP = "驳回任务";
            stepModel.STATUS = model.Status;
            stepModel.UpdateEmpNo = model.EmpNo;
            stepModel.UpdateEmpName = model.FullName;
            stepModel.UpdateTime = DateTime.Now;

            if (model.Status == "complete" || model.Status == "rejectend")
            {
                model.FinishTime = DateTime.Now;
            }
            try
            {
                string sql = "";
                using (RepairDb db = new RepairDb())
                {
                    sql = @" update [888_KsNorth].[dbo].[record]  set ReplyContent = concat(ReplyContent,@NewReplyContent), Status = @Status,FinishTime = @FinishTime,charge_empno = @ChargeEmpno,charge_empname = @ChargeEmpname where repair_id = @RepairId ";

                    string stepSQL = @"insert into [888_KsNorth].[dbo].[steprecord] ([guid],[repair_id],[status],[OPINION],[step],[charge_empno],[charge_empname],[UpdateEmpNo],[UpdateEmpName],[UpdateTime]) 
                                            values (@GUID,@RepairId,@STATUS,@OPINION,@STEP,@ChargeEmpno,@ChargeEmpname,@UpdateEmpNo,@UpdateEmpName,@UpdateTime)";

                    //判断是否为转派单，否则无法驳回任务
                    string checkTransSql = @" select * from [888_KsNorth].[dbo].[steprecord] where step = '转派报修单' and charge_empno =@ChargeEmpno and repair_id = @RepairId ";
                    var checkModel = db.Query<StepRecordModel>(checkTransSql, new { ChargeEmpno = model.ChargeEmpno, RepairId = model.RepairId }).ToList();
                    if (checkModel.Count() == 0)
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "此单并非为转派单，无法驳回任务" });
                    }
                    //抓取原来的负责人
                    string findCharge = @" SELECT a.*,a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname FROM [888_KsNorth].[dbo].[steprecord] a WHERE a.repair_id = @RepairId
                                             AND a.sort < (select TOP(1) sort from [888_KsNorth].[dbo].[steprecord] where repair_id = @RepairId and STEP = '转派报修单' AND charge_empno = @ChargeEmpno order by sort desc )  order by sort desc ";
                    var chargeModel = db.Query<StepRecordModel>(findCharge, new { RepairId = model.RepairId, ChargeEmpno = model.ChargeEmpno }).FirstOrDefault();

                    //抓取原来的转派人
                    string findTransfer = @" SELECT a.*,a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname FROM [888_KsNorth].[dbo].[steprecord] a WHERE a.repair_id = @RepairId
                                             AND a.sort = (select TOP(1) sort from [888_KsNorth].[dbo].[steprecord] where repair_id = @RepairId and STEP = '转派报修单' order by sort desc )  order by sort desc ";
                    var transferModel = db.Query<StepRecordModel>(findTransfer, new { RepairId = model.RepairId }).FirstOrDefault();

                    stepModel.ChargeEmpno = chargeModel.ChargeEmpno;
                    stepModel.ChargeEmpname = chargeModel.ChargeEmpname;

                    model.ChargeEmpno = chargeModel.ChargeEmpno;
                    model.ChargeEmpname = chargeModel.ChargeEmpname;

                    model.NewReplyContent = model.NewReplyContent + "\r\n-" + model.FullName + "于" + DateTime.Now + "驳回任务给-" + model.ChargeEmpname + "\r\n\r\n";

                    SendMailService sendMailService = new SendMailService();
                    //报修单号  原来的负责人  原来转派的人 
                    bool mailFlag = sendMailService.rejectMail(model.RepairId, chargeModel.ChargeEmpno, transferModel.UpdateEmpNo, model.FullName);
                    if (mailFlag)
                    {
                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(sql, model);
                        trans.Add(stepSQL, stepModel);
                        db.DoExtremeSpeedTransaction(trans);
                    }
                    else
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "邮件通知异常，驳回任务失败" });
                    }

                }
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
            return Json(new FlagTips { IsSuccess = true });
        }

        /// <summary>
        /// 报修单 驳回至使用者
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult RejectEnd(RepairRecordModel model)
        {
            model.EmpNo = Session["EmpNo"].ToString();
            model.FullName = Session["fullname"].ToString();

            StepRecordModel stepModel = new StepRecordModel();
            stepModel.OPINION = model.NewReplyContent;//获取未重新定义前的新回复内容
            stepModel.RepairId = model.RepairId;
            stepModel.STATUS = "rejectend";
            stepModel.STEP = "驳回至使用者";
            stepModel.ChargeEmpno = model.ChargeEmpno;
            stepModel.ChargeEmpname = model.ChargeEmpname;
            stepModel.UpdateEmpNo = model.EmpNo;
            stepModel.UpdateEmpName = model.FullName;
            stepModel.UpdateTime = DateTime.Now;
            stepModel.GUID = Guid.NewGuid().ToString();

            model.NewReplyContent = model.NewReplyContent + "\r\n-" + model.FullName + "于" + DateTime.Now + "驳回给反应者-" + model.ResponseEmpname + "\r\n\r\n";
            model.Status = "rejectend";
            model.FinishTime = DateTime.Now;
            try
            {
                string sql = "";
                model.ResponseEmpno = Session["EmpNo"].ToString();
                model.ResponseEmpname = Session["fullname"].ToString();
                model.CreatTime = DateTime.Now;
                using (RepairDb db = new RepairDb())
                {
                    sql = @" update [888_KsNorth].[dbo].[record]  set ReplyContent = concat(ReplyContent,@NewReplyContent),Status = @Status,FinishTime = @FinishTime where repair_id = @RepairId ";

                    string stepSQL = @"insert into [888_KsNorth].[dbo].[steprecord] ([guid],[repair_id],[status],[step],[OPINION],[charge_empno],[charge_empname],[UpdateEmpNo],[UpdateEmpName],[UpdateTime]) 
                                            values (@GUID,@RepairId,@STATUS,@STEP,@OPINION,@ChargeEmpno,@ChargeEmpname,@UpdateEmpNo,@UpdateEmpName,@UpdateTime)";

                    SendMailService sendMailService = new SendMailService();
                    bool mailFlag = sendMailService.rejectEndMail(model.RepairId);
                    if (mailFlag)
                    {
                        Dictionary<string, object> trans = new Dictionary<string, object>();
                        trans.Add(sql, model);
                        trans.Add(stepSQL, stepModel);
                        db.DoExtremeSpeedTransaction(trans);
                    }
                    else
                    {
                        return Json(new FlagTips { IsSuccess = false, Msg = "邮件通知异常，驳回任务失败" });
                    }

                }
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
            return Json(new FlagTips { IsSuccess = true });
        }
        #endregion

        #region 问题列表

        /// <summary>
        /// 问题列表VW
        /// </summary>
        /// <returns></returns>
        public ActionResult ProblemListIndex()
        {
            return View();
        }


        /// <summary>
        /// 问题列表查询
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult getProblemList(RepairRecordModel model)
        {
            var list = new List<RepairRecordModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT DISTINCT a.repair_id RepairId,a.area_id AreaId,a.kind_id KindId,a.SystemCategory,a.Building Building,a.Loaction,
                                               a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname,a.Category,a.ResponseContent,a.ReplyContent,b.StatusText Status,a.CreatTime,
                                               a.PhotoPath,RoomNum,a.repairTime RepairTime,a.Telephone,a.DamageReason,a.DamageClass,a.DamageName,a.ResponseEmpno,
                                               a.ResponseEmpname,a.FinishTime from [888_KsNorth].[dbo].[record] a 
                                               left join [888_KsNorth].[dbo].[state] b on a.status = b.StatusValue and a.SystemCategory = b.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] area  ON a.area_id = area.area_id AND area.match_type = 'AreaMatch' AND a.SystemCategory = area.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] kind  ON a.kind_id = kind.area_id AND kind.match_type = 'KindMatch' AND a.SystemCategory = kind.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch1  ON area.charge_emp = ch1.EmpNo AND a.SystemCategory = ch1.SystemCategory 
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch2  ON area.charge_emp = ch2.EmpNo AND a.SystemCategory = ch2.SystemCategory WHERE 1= 1
											  ");
                    //系统类别
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and a.SystemCategory =@SystemCategory ";
                    }
                    //大楼别
                    if (!string.IsNullOrEmpty(model.Building))
                    {
                        sql += " and a.Building =@Building ";
                    }
                    //位置
                    if (!string.IsNullOrEmpty(model.Loaction))
                    {
                        sql += " and a.Loaction =@Loaction ";
                    }
                    //类别
                    if (!string.IsNullOrEmpty(model.Category))
                    {
                        sql += " and a.Category =@Category ";
                    }
                    //反应人姓名
                    if (!string.IsNullOrEmpty(model.ResponseEmpname))
                    {
                        sql += " and a.ResponseEmpname like '%" + model.ResponseEmpname + "%'";
                    }
                    //反应内容
                    if (!string.IsNullOrEmpty(model.ResponseContent))
                    {
                        sql += " and a.ResponseContent like '%" + model.ResponseContent + "%'";
                    }
                    //执行状态
                    if (!string.IsNullOrEmpty(model.Status))
                    {
                        sql += " and a.Status =@Status ";
                    }
                    //负责人
                    if (!string.IsNullOrEmpty(model.ChargeEmpname))
                    {
                        sql += " and a.charge_empname =@ChargeEmpname ";
                    }
                    //填表时间
                    if (model.startDate != null)
                    {
                        sql += " and a.CreatTime >= @startDate ";
                    }
                    if (model.endDate != null)
                    {
                        model.endDate = Convert.ToDateTime(model.endDate).AddDays(1);
                        sql += " and a.CreatTime <=@endDate ";
                    }
                    var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
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

        /// <summary>
        /// 问题列表 报修单详细信息
        /// </summary>
        /// <returns></returns>
        public ActionResult ProblemDetailVw()
        {
            return View();
        }


        public List<RepairRecordModel> queryProblemList(RepairRecordModel model)
        {
            var list = new List<RepairRecordModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT DISTINCT a.repair_id RepairId,a.area_id AreaId,a.kind_id KindId,a.SystemCategory,a.Building Building,a.Loaction,
                                               a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname,a.Category,a.ResponseContent,a.ReplyContent,b.StatusText Status,a.CreatTime,
                                               a.PhotoPath,RoomNum,a.repairTime RepairTime,a.Telephone,a.DamageReason,a.DamageClass,a.DamageName,a.ResponseEmpno,
                                               a.ResponseEmpname,a.FinishTime from [888_KsNorth].[dbo].[record] a 
                                               left join [888_KsNorth].[dbo].[state] b on a.status = b.StatusValue and a.SystemCategory = b.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] area  ON a.area_id = area.area_id AND area.match_type = 'AreaMatch' AND a.SystemCategory = area.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] kind  ON a.kind_id = kind.area_id AND kind.match_type = 'KindMatch' AND a.SystemCategory = kind.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch1  ON area.charge_emp = ch1.EmpNo AND a.SystemCategory = ch1.SystemCategory 
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch2  ON area.charge_emp = ch2.EmpNo AND a.SystemCategory = ch2.SystemCategory WHERE 1= 1
											  ");
                    //系统类别
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and a.SystemCategory =@SystemCategory ";
                    }
                    //大楼别
                    if (!string.IsNullOrEmpty(model.Building))
                    {
                        sql += " and a.Building =@Building ";
                    }
                    //位置
                    if (!string.IsNullOrEmpty(model.Loaction))
                    {
                        sql += " and a.Loaction =@Loaction ";
                    }
                    //类别
                    if (!string.IsNullOrEmpty(model.Category))
                    {
                        sql += " and a.Category =@Category ";
                    }
                    //反应人姓名
                    if (!string.IsNullOrEmpty(model.ResponseEmpname))
                    {
                        sql += " and a.ResponseEmpname like '%" + model.ResponseEmpname + "%'";
                    }
                    //反应内容
                    if (!string.IsNullOrEmpty(model.ResponseContent))
                    {
                        sql += " and a.ResponseContent like '%" + model.ResponseContent + "%'";
                    }
                    //执行状态
                    if (!string.IsNullOrEmpty(model.Status))
                    {
                        sql += " and a.Status =@Status ";
                    }
                    //负责人
                    if (!string.IsNullOrEmpty(model.ChargeEmpname))
                    {
                        sql += " and a.charge_empname =@ChargeEmpname ";
                    }
                    //填表时间
                    if (model.startDate != null)
                    {
                        sql += " and a.CreatTime >= @startDate ";
                    }
                    if (model.endDate != null)
                    {
                        model.endDate = Convert.ToDateTime(model.endDate).AddDays(1);
                        sql += " and a.CreatTime <=@endDate ";
                    }
                    var group = Session["GroupName"] == null ? Session["OPGroup"].ToString() : Session["GroupName"].ToString();
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
        public ActionResult ProblemListExport(RepairRecordModel model)
        {
            try
            {
                var dt = queryProblemList(model);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string templatePath = string.Format("~\\Excel\\ProblemList.xlsx");
                FileStream fs = new FileStream(System.Web.HttpContext.Current.Server.MapPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                Workbook wb = new Workbook(fs);
                Worksheet sheet = wb.Worksheets[0];
                sheet.Name = "问题列表";
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
                    sheet.Cells[i + 1, 0].PutValue(dt[i].RepairId);
                    sheet.Cells[i + 1, 1].PutValue(dt[i].SystemCategory);
                    sheet.Cells[i + 1, 2].PutValue(dt[i].Building);
                    sheet.Cells[i + 1, 3].PutValue(dt[i].Loaction);
                    sheet.Cells[i + 1, 4].PutValue(dt[i].Category);
                    sheet.Cells[i + 1, 5].PutValue(dt[i].ResponseEmpname);
                    sheet.Cells[i + 1, 6].PutValue(dt[i].ChargeEmpname);
                    sheet.Cells[i + 1, 7].PutValue(dt[i].ResponseContent);
                    sheet.Cells[i + 1, 8].PutValue(dt[i].Status);
                    sheet.Cells[i + 1, 9].PutValue(dt[i].CreatTime.ToString() == "0001/01/01" ? "" : dt[i].CreatTime.ToString());
                    sheet.Cells[i + 1, 10].PutValue(dt[i].FinishTime.ToString() == "0001/01/01" ? "" : dt[i].FinishTime.ToString());
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

        /// <summary>
        /// 步骤记录
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult getStepList(string RepairId)
        {
            var list = new List<StepRecordModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@" select a.*,a.repair_id RepairId,a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname  from [888_KsNorth].[dbo].[steprecord]  a
											       WHERE a.repair_id= @RepairId ORDER BY a.sort asc");

                    list = db.Query<StepRecordModel>(sql, new { RepairId }).ToList();
                }
                return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        #endregion

        #region  问题列表维护
        /// <summary>
        /// 问题列表维护
        /// </summary>
        /// <returns></returns>
        public ActionResult ProblemAllVw()
        {
            return View();
        }

        /// <summary>
        /// 问题列表维护查询
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public ActionResult getProblemAllList(RepairRecordModel model)
        {
            var list = new List<RepairRecordModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = string.Format(@"SELECT DISTINCT a.repair_id RepairId,a.area_id AreaId,a.kind_id KindId,a.SystemCategory,a.Building Building,a.Loaction,
                                               a.charge_empno ChargeEmpno,a.charge_empname ChargeEmpname,a.Category,a.ResponseContent,a.ReplyContent,b.StatusText Status,a.CreatTime,
                                               a.PhotoPath,RoomNum,a.repairTime RepairTime,a.Telephone,a.DamageReason,a.DamageClass,a.DamageName,a.ResponseEmpno,
                                               a.ResponseEmpname,a.FinishTime from [888_KsNorth].[dbo].[record] a 
                                               left join [888_KsNorth].[dbo].[state] b on a.status = b.StatusValue and a.SystemCategory = b.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] area  ON a.area_id = area.area_id AND area.match_type = 'AreaMatch' AND a.SystemCategory = area.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[match] kind  ON a.kind_id = kind.area_id AND kind.match_type = 'KindMatch' AND a.SystemCategory = kind.SystemCategory
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch1  ON area.charge_emp = ch1.EmpNo AND a.SystemCategory = ch1.SystemCategory 
											   LEFT JOIN [888_KsNorth].[dbo].[charge] ch2  ON area.charge_emp = ch2.EmpNo AND a.SystemCategory = ch2.SystemCategory WHERE 1= 1
											  ");
                    //系统类别
                    if (!string.IsNullOrEmpty(model.SystemCategory))
                    {
                        sql += " and a.SystemCategory =@SystemCategory ";
                    }
                    //大楼别
                    if (!string.IsNullOrEmpty(model.Building))
                    {
                        sql += " and a.Building =@Building ";
                    }
                    //位置
                    if (!string.IsNullOrEmpty(model.Loaction))
                    {
                        sql += " and a.Loaction =@Loaction ";
                    }
                    //类别
                    if (!string.IsNullOrEmpty(model.Category))
                    {
                        sql += " and a.Category =@Category ";
                    }
                    //反应人姓名
                    if (!string.IsNullOrEmpty(model.ResponseEmpname))
                    {
                        sql += " and a.ResponseEmpname like '%" + model.ResponseEmpname + "%'";
                    }
                    //反应内容
                    if (!string.IsNullOrEmpty(model.ResponseContent))
                    {
                        sql += " and a.ResponseContent like '%" + model.ResponseContent + "%'";
                    }
                    //执行状态
                    if (!string.IsNullOrEmpty(model.Status))
                    {
                        sql += " and a.Status =@Status ";
                    }
                    //负责人
                    if (!string.IsNullOrEmpty(model.ChargeEmpname))
                    {
                        sql += " and a.charge_empname =@ChargeEmpname ";
                    }
                    //填表时间
                    if (model.startDate != null)
                    {
                        sql += " and a.CreatTime >= @startDate ";
                    }
                    if (model.endDate != null)
                    {
                        model.endDate = Convert.ToDateTime(model.endDate).AddDays(1);
                        sql += " and a.CreatTime <=@endDate ";
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

        /// <summary>
        /// 问题列表维护 报修单详情
        /// </summary>
        /// <returns></returns>
        public ActionResult ProblemAllDetailVw()
        {
            return View();
        }


        #endregion
    }
}