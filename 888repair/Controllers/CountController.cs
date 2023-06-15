using _888repair.Db;
using _888repair.Models;
using _888repair.Models.Count;
using Aspose.Cells;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace _888repair.Controllers
{
    public class CountController : Controller
    {
        // GET: Count  统计页面
        public ActionResult CountIndex()
        {
            return View();
        }

        public ActionResult getCountList(QueryCountModel model)
        {
            try
            {
                var list = new List<CountModel>();
                if (string.IsNullOrEmpty(model.QueryType))
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = "请先选择统计类别" }, JsonRequestBehavior.AllowGet);
                }
                if (model.StartDate == null)
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = "请先选择起始日期" }, JsonRequestBehavior.AllowGet);
                }
                if (model.EndDate == null)
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = "请先选择终止日期" }, JsonRequestBehavior.AllowGet);
                }
                else if (model.EndDate != null)
                {
                    model.EndDate = Convert.ToDateTime(model.EndDate).AddDays(1);
                }
                using (RepairDb db = new RepairDb())
                {
                    string sql = "";
                    switch (model.QueryType)
                    {
                        #region
                        case "department"://部门统计
                            sql = @" SELECT SUM(   CASE a.Status
                  WHEN 'complete' THEN
                      1
                  WHEN 'rejectend' THEN
                      1
                  ELSE
                      0
              END
          ) AS FinishCount,
       COUNT(*) AS TotalCount,
       CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
       COUNT(*) - SUM(   CASE a.Status
                             WHEN 'complete' THEN
                                 1
                             WHEN 'rejectend' THEN
                                 1
                             ELSE
                                 0
                         END
                     ) AS NotFinishCount,
       c.DeptName Department
FROM [888_KsNorth].[dbo].[record] a 
LEFT JOIN[Common].[dbo].[kcis_account] B
ON A.ResponseEmpno = b.EmpNo
LEFT JOIN [Common].[dbo].[AFS_Dept] c
        ON b.deptid2 = c.DeptID_eip
WHERE 1=1 
      AND a.SystemCategory = @SystemCategory
	  AND a.CreatTime>= @StartDate
	  AND a.CreatTime <= @EndDate
GROUP BY c.DeptName";
                            break;
                        case "charge"://负责人统计
                            sql = @"
SELECT S.FullName ChargeName,
      ISNULL(g.TotalCount,'0')  TotalCount,
	  ISNULL(g.FinishCount,'0')  FinishCount,
	  ISNULL(g.NotFinishCount,'0')  NotFinishCount,
	  ISNULL(g.CompletionRate,'0%')  CompletionRate
FROM [888_KsNorth].[dbo].[charge] S
    LEFT JOIN
    (
        SELECT SUM(   CASE a.Status
                          WHEN 'complete' THEN
                              1
                          WHEN 'rejectend' THEN
                              1
                          ELSE
                              0
                      END
                  ) AS FinishCount,
               COUNT(*) AS TotalCount,
               CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
               COUNT(*) - SUM(   CASE a.Status
                                     WHEN 'complete' THEN
                                         1
                                     WHEN 'rejectend' THEN
                                         1
                                     ELSE
                                         0
                                 END
                             ) AS NotFinishCount,
               a.charge_empno,
               a.SystemCategory
        FROM [888_KsNorth].[dbo].[record] a
        WHERE 1 = 1
              AND a.SystemCategory = @SystemCategory
			  AND a.CreatTime>= @StartDate
			  AND a.CreatTime <= @EndDate
        GROUP BY a.charge_empno,
                 a.SystemCategory
    ) G
        ON S.EmpNo = G.charge_empno
           AND S.SystemCategory = G.SystemCategory
		   WHERE s.SystemCategory = @SystemCategory ";
                            break;
                        case "response"://反应人统计
                            sql = @"SELECT SUM(   CASE a.Status
                  WHEN 'complete' THEN
                      1
                  WHEN 'rejectend' THEN
                      1
                  ELSE
                      0
              END
          ) AS FinishCount,
       COUNT(*) AS TotalCount,
       CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
       COUNT(*) - SUM(   CASE a.Status
                             WHEN 'complete' THEN
                                 1
                             WHEN 'rejectend' THEN
                                 1
                             ELSE
                                 0
                         END
                     ) AS NotFinishCount,
       a.ResponseEmpname ResponseName
FROM [888_KsNorth].[dbo].[record] a
WHERE 1=1 
      AND a.SystemCategory = @SystemCategory
	  AND a.CreatTime>= @StartDate
	  AND a.CreatTime <= @EndDate
GROUP BY a.ResponseEmpname";
                            break;
                        case "area"://辖区统计
                            sql = @"SELECT S.Buliding Building,
       S.Location,
       E.FullName ChargeName,
       ISNULL(G.TotalCount, '0') TotalCount,
       ISNULL(G.FinishCount, '0') FinishCount,
       ISNULL(G.NotFinishCount, '0') NotFinishCount,
       ISNULL(G.CompletionRate, '0%') CompletionRate
FROM [888_KsNorth].[dbo].[area] S
    LEFT JOIN
    (
        SELECT SUM(   CASE a.Status
                          WHEN 'complete' THEN
                              1
                          WHEN 'rejectend' THEN
                              1
                          ELSE
                              0
                      END
                  ) AS FinishCount,
               COUNT(*) AS TotalCount,
               CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
               COUNT(*) - SUM(   CASE a.Status
                                     WHEN 'complete' THEN
                                         1
                                     WHEN 'rejectend' THEN
                                         1
                                     ELSE
                                         0
                                 END
                             ) AS NotFinishCount,
               a.area_id,
               a.SystemCategory
        FROM [888_KsNorth].[dbo].[record] a
        WHERE 1 = 1
              AND a.SystemCategory = @SystemCategory
			  AND a.CreatTime>= @StartDate
			  AND a.CreatTime <= @EndDate
        GROUP BY a.area_id,
                 a.SystemCategory
    ) G
        ON S.area_id = G.area_id
           AND S.SystemCategory = G.SystemCategory
    LEFT JOIN [888_KsNorth].[dbo].[match] M
        ON S.area_id = M.area_id
           AND M.SystemCategory = S.SystemCategory
           AND M.match_type = 'AreaMatch'
    LEFT JOIN [888_KsNorth].[dbo].[charge] E
        ON S.SystemCategory = E.SystemCategory
           AND M.charge_emp = E.EmpNo
WHERE S.SystemCategory = @SystemCategory
      AND M.sort = 1
ORDER BY S.area_id ASC";
                            break;
                        case "kind"://业管统计
                            sql = @"SELECT S.KindCategory Category,
       E.FullName ChargeName,
       ISNULL(G.TotalCount, '0') TotalCount,
       ISNULL(G.FinishCount, '0') FinishCount,
       ISNULL(G.NotFinishCount, '0') NotFinishCount,
       ISNULL(G.CompletionRate, '0%') CompletionRate
FROM [888_KsNorth].[dbo].[kind] S
    LEFT JOIN
    (
        SELECT SUM(   CASE a.Status
                          WHEN 'complete' THEN
                              1
                          WHEN 'rejectend' THEN
                              1
                          ELSE
                              0
                      END
                  ) AS FinishCount,
               COUNT(*) AS TotalCount,
               CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
               COUNT(*) - SUM(   CASE a.Status
                                     WHEN 'complete' THEN
                                         1
                                     WHEN 'rejectend' THEN
                                         1
                                     ELSE
                                         0
                                 END
                             ) AS NotFinishCount,
               a.kind_id,
               a.SystemCategory
        FROM [888_KsNorth].[dbo].[record] a
        WHERE 1 = 1
              AND a.SystemCategory = @SystemCategory
			  AND a.CreatTime>= @StartDate
			  AND a.CreatTime <= @EndDate
        GROUP BY a.kind_id,
                 a.SystemCategory
    ) G
        ON S.kind_id = G.kind_id
           AND S.SystemCategory = G.SystemCategory
    LEFT JOIN [888_KsNorth].[dbo].[match] M
        ON S.kind_id = M.area_id
           AND M.SystemCategory = S.SystemCategory
           AND M.match_type = 'KindMatch'
    LEFT JOIN [888_KsNorth].[dbo].[charge] E
        ON S.SystemCategory = E.SystemCategory
           AND M.charge_emp = E.EmpNo
WHERE S.SystemCategory = @SystemCategory
      AND M.sort = 1 
ORDER BY S.kind_id ASC ";
                            break;
                        default:
                            break;
                            #endregion
                    }

                    list = db.Query<CountModel>(sql, model).ToList();
                }
                    return Json(new FlagTips { IsSuccess = true, code = 0, count = list.Count(), data = list }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public List<CountModel> getCountListToExcel(QueryCountModel model)
        {
            var list = new List<CountModel>();
            try
            {
                using (RepairDb db = new RepairDb())
                {
                    string sql = "";
                    switch (model.QueryType)
                    {
                        #region
                        case "department"://部门统计
                            sql = @" SELECT SUM(   CASE a.Status
                  WHEN 'complete' THEN
                      1
                  WHEN 'rejectend' THEN
                      1
                  ELSE
                      0
              END
          ) AS FinishCount,
       COUNT(*) AS TotalCount,
       CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
       COUNT(*) - SUM(   CASE a.Status
                             WHEN 'complete' THEN
                                 1
                             WHEN 'rejectend' THEN
                                 1
                             ELSE
                                 0
                         END
                     ) AS NotFinishCount,
       c.DeptName Department
FROM [888_KsNorth].[dbo].[record] a 
LEFT JOIN[Common].[dbo].[kcis_account] B
ON A.ResponseEmpno = b.EmpNo
LEFT JOIN [Common].[dbo].[AFS_Dept] c
        ON b.deptid2 = c.DeptID_eip
WHERE 1=1 
      AND a.SystemCategory = @SystemCategory
	  AND a.CreatTime>= @StartDate
	  AND a.CreatTime <= @EndDate
GROUP BY c.DeptName";
                            break;
                        case "charge"://负责人统计
                            sql = @"
SELECT S.FullName ChargeName,
      ISNULL(g.TotalCount,'0')  TotalCount,
	  ISNULL(g.FinishCount,'0')  FinishCount,
	  ISNULL(g.NotFinishCount,'0')  NotFinishCount,
	  ISNULL(g.CompletionRate,'0%')  CompletionRate
FROM [888_KsNorth].[dbo].[charge] S
    LEFT JOIN
    (
        SELECT SUM(   CASE a.Status
                          WHEN 'complete' THEN
                              1
                          WHEN 'rejectend' THEN
                              1
                          ELSE
                              0
                      END
                  ) AS FinishCount,
               COUNT(*) AS TotalCount,
               CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
               COUNT(*) - SUM(   CASE a.Status
                                     WHEN 'complete' THEN
                                         1
                                     WHEN 'rejectend' THEN
                                         1
                                     ELSE
                                         0
                                 END
                             ) AS NotFinishCount,
               a.charge_empno,
               a.SystemCategory
        FROM [888_KsNorth].[dbo].[record] a
        WHERE 1 = 1
              AND a.SystemCategory = @SystemCategory
			  AND a.CreatTime>= @StartDate
			  AND a.CreatTime <= @EndDate
        GROUP BY a.charge_empno,
                 a.SystemCategory
    ) G
        ON S.EmpNo = G.charge_empno
           AND S.SystemCategory = G.SystemCategory
		   WHERE s.SystemCategory = @SystemCategory ";
                            break;
                        case "response"://反应人统计
                            sql = @"SELECT SUM(   CASE a.Status
                  WHEN 'complete' THEN
                      1
                  WHEN 'rejectend' THEN
                      1
                  ELSE
                      0
              END
          ) AS FinishCount,
       COUNT(*) AS TotalCount,
       CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
       COUNT(*) - SUM(   CASE a.Status
                             WHEN 'complete' THEN
                                 1
                             WHEN 'rejectend' THEN
                                 1
                             ELSE
                                 0
                         END
                     ) AS NotFinishCount,
       a.ResponseEmpname ResponseName
FROM [888_KsNorth].[dbo].[record] a
WHERE 1=1 
      AND a.SystemCategory = @SystemCategory
	  AND a.CreatTime>= @StartDate
	  AND a.CreatTime <= @EndDate
GROUP BY a.ResponseEmpname";
                            break;
                        case "area"://辖区统计
                            sql = @"SELECT S.Buliding Building,
       S.Location,
       E.FullName ChargeName,
       ISNULL(G.TotalCount, '0') TotalCount,
       ISNULL(G.FinishCount, '0') FinishCount,
       ISNULL(G.NotFinishCount, '0') NotFinishCount,
       ISNULL(G.CompletionRate, '0%') CompletionRate
FROM [888_KsNorth].[dbo].[area] S
    LEFT JOIN
    (
        SELECT SUM(   CASE a.Status
                          WHEN 'complete' THEN
                              1
                          WHEN 'rejectend' THEN
                              1
                          ELSE
                              0
                      END
                  ) AS FinishCount,
               COUNT(*) AS TotalCount,
               CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
               COUNT(*) - SUM(   CASE a.Status
                                     WHEN 'complete' THEN
                                         1
                                     WHEN 'rejectend' THEN
                                         1
                                     ELSE
                                         0
                                 END
                             ) AS NotFinishCount,
               a.area_id,
               a.SystemCategory
        FROM [888_KsNorth].[dbo].[record] a
        WHERE 1 = 1
              AND a.SystemCategory = @SystemCategory
			  AND a.CreatTime>= @StartDate
			  AND a.CreatTime <= @EndDate
        GROUP BY a.area_id,
                 a.SystemCategory
    ) G
        ON S.area_id = G.area_id
           AND S.SystemCategory = G.SystemCategory
    LEFT JOIN [888_KsNorth].[dbo].[match] M
        ON S.area_id = M.area_id
           AND M.SystemCategory = S.SystemCategory
           AND M.match_type = 'AreaMatch'
    LEFT JOIN [888_KsNorth].[dbo].[charge] E
        ON S.SystemCategory = E.SystemCategory
           AND M.charge_emp = E.EmpNo
WHERE S.SystemCategory = @SystemCategory
      AND M.sort = 1
ORDER BY S.area_id ASC";
                            break;
                        case "kind"://业管统计
                            sql = @"SELECT S.KindCategory Category,
       E.FullName ChargeName,
       ISNULL(G.TotalCount, '0') TotalCount,
       ISNULL(G.FinishCount, '0') FinishCount,
       ISNULL(G.NotFinishCount, '0') NotFinishCount,
       ISNULL(G.CompletionRate, '0%') CompletionRate
FROM [888_KsNorth].[dbo].[kind] S
    LEFT JOIN
    (
        SELECT SUM(   CASE a.Status
                          WHEN 'complete' THEN
                              1
                          WHEN 'rejectend' THEN
                              1
                          ELSE
                              0
                      END
                  ) AS FinishCount,
               COUNT(*) AS TotalCount,
               CONCAT(CONVERT(DECIMAL(18, 2),
                              CAST(SUM(   CASE a.Status
                                              WHEN 'complete' THEN
                                                  1
                                              WHEN 'rejectend' THEN
                                                  1
                                              ELSE
                                                  0
                                          END
                                      ) AS FLOAT) / CAST(COUNT(*) AS FLOAT) * 100
                             ),
                      '%'
                     ) AS CompletionRate,
               COUNT(*) - SUM(   CASE a.Status
                                     WHEN 'complete' THEN
                                         1
                                     WHEN 'rejectend' THEN
                                         1
                                     ELSE
                                         0
                                 END
                             ) AS NotFinishCount,
               a.kind_id,
               a.SystemCategory
        FROM [888_KsNorth].[dbo].[record] a
        WHERE 1 = 1
              AND a.SystemCategory = @SystemCategory
			  AND a.CreatTime>= @StartDate
			  AND a.CreatTime <= @EndDate
        GROUP BY a.kind_id,
                 a.SystemCategory
    ) G
        ON S.kind_id = G.kind_id
           AND S.SystemCategory = G.SystemCategory
    LEFT JOIN [888_KsNorth].[dbo].[match] M
        ON S.kind_id = M.area_id
           AND M.SystemCategory = S.SystemCategory
           AND M.match_type = 'KindMatch'
    LEFT JOIN [888_KsNorth].[dbo].[charge] E
        ON S.SystemCategory = E.SystemCategory
           AND M.charge_emp = E.EmpNo
WHERE S.SystemCategory = @SystemCategory
      AND M.sort = 1 
ORDER BY S.kind_id ASC ";
                            break;
                        default:
                            break;
                            #endregion
                    }
                    list = db.Query<CountModel>(sql, model).ToList();
                }
            }
            catch (Exception ex)
            {

            }
            return list;
        }

        public ActionResult CountExport(QueryCountModel model)
        {
            try
            {
                if (string.IsNullOrEmpty(model.QueryType))
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = "请先选择统计类别" }, JsonRequestBehavior.AllowGet);
                }
                if (model.StartDate == null)
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = "请先选择起始日期" }, JsonRequestBehavior.AllowGet);
                }
                if (model.EndDate == null)
                {
                    return Json(new FlagTips { IsSuccess = false, Msg = "请先选择终止日期" }, JsonRequestBehavior.AllowGet);
                }
                else if (model.EndDate != null)
                {
                    model.EndDate = Convert.ToDateTime(model.EndDate).AddDays(1);
                }
                var dt = getCountListToExcel(model);
                //部门统计
                MemoryStream bookStream = new MemoryStream();
                string fileName = "";
                if (model.QueryType == "department")
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    string templatePath = string.Format("~\\Excel\\DepartmentCount.xlsx");
                    FileStream fs = new FileStream(System.Web.HttpContext.Current.Server.MapPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    Workbook wb = new Workbook(fs);
                    Worksheet sheet = wb.Worksheets[0];
                    sheet.Name = "部门统计";
                    fileName = sheet.Name;
                    Cells cells = sheet.Cells;
                    int columnCount = cells.MaxColumn;
                    int rowCount = cells.MaxRow; 

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
                        sheet.Cells[i + 1, 0].PutValue(dt[i].Department);
                        sheet.Cells[i + 1, 1].PutValue(dt[i].TotalCount);
                        sheet.Cells[i + 1, 2].PutValue(dt[i].FinishCount);
                        sheet.Cells[i + 1, 3].PutValue(dt[i].NotFinishCount);
                        sheet.Cells[i + 1, 4].PutValue(dt[i].CompletionRate);
                    }
                  //  MemoryStream bookStream = new MemoryStream();
                    wb.Save(bookStream, new OoxmlSaveOptions(SaveFormat.Xlsx));
                    bookStream.Seek(0, SeekOrigin.Begin);
                  //  return File(bookStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("部门统计_{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff")));//最后以文件形式返回
                }
                //负责人统计
                else if (model.QueryType == "charge")
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    string templatePath = string.Format("~\\Excel\\ChargeCount.xlsx");
                    FileStream fs = new FileStream(System.Web.HttpContext.Current.Server.MapPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    Workbook wb = new Workbook(fs);
                    Worksheet sheet = wb.Worksheets[0];
                    sheet.Name = "负责人统计";
                    fileName = sheet.Name;
                    Cells cells = sheet.Cells;
                    int columnCount = cells.MaxColumn;
                    int rowCount = cells.MaxRow;

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
                        sheet.Cells[i + 1, 0].PutValue(dt[i].ChargeName);
                        sheet.Cells[i + 1, 1].PutValue(dt[i].TotalCount);
                        sheet.Cells[i + 1, 2].PutValue(dt[i].FinishCount);
                        sheet.Cells[i + 1, 3].PutValue(dt[i].NotFinishCount);
                        sheet.Cells[i + 1, 4].PutValue(dt[i].CompletionRate);
                    }
                 //   MemoryStream bookStream = new MemoryStream();
                    wb.Save(bookStream, new OoxmlSaveOptions(SaveFormat.Xlsx));
                    bookStream.Seek(0, SeekOrigin.Begin);
                   // return File(bookStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("负责人统计_{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff")));//最后以文件形式返回
                }
                //反应人统计 response
                else if (model.QueryType == "response")
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    string templatePath = string.Format("~\\Excel\\ResponseCount.xlsx");
                    FileStream fs = new FileStream(System.Web.HttpContext.Current.Server.MapPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    Workbook wb = new Workbook(fs);
                    Worksheet sheet = wb.Worksheets[0];
                    sheet.Name = "反应人统计";
                    fileName = sheet.Name;
                    Cells cells = sheet.Cells;
                    int columnCount = cells.MaxColumn;
                    int rowCount = cells.MaxRow;

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
                        sheet.Cells[i + 1, 0].PutValue(dt[i].ResponseName);
                        sheet.Cells[i + 1, 1].PutValue(dt[i].TotalCount);
                        sheet.Cells[i + 1, 2].PutValue(dt[i].FinishCount);
                        sheet.Cells[i + 1, 3].PutValue(dt[i].NotFinishCount);
                        sheet.Cells[i + 1, 4].PutValue(dt[i].CompletionRate);
                    }
                  //  MemoryStream bookStream = new MemoryStream();
                    wb.Save(bookStream, new OoxmlSaveOptions(SaveFormat.Xlsx));
                    bookStream.Seek(0, SeekOrigin.Begin);
                  //  return File(bookStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("反应人统计_{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff")));//最后以文件形式返回
                }
                //辖区统计 area
                else if (model.QueryType == "area")
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    string templatePath = string.Format("~\\Excel\\AreaCount.xlsx");
                    FileStream fs = new FileStream(System.Web.HttpContext.Current.Server.MapPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    Workbook wb = new Workbook(fs);
                    Worksheet sheet = wb.Worksheets[0];
                    sheet.Name = "辖区统计";
                    fileName = sheet.Name;
                    Cells cells = sheet.Cells;
                    int columnCount = cells.MaxColumn;
                    int rowCount = cells.MaxRow;

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
                        sheet.Cells[i + 1, 0].PutValue(dt[i].Building);
                        sheet.Cells[i + 1, 1].PutValue(dt[i].Location);
                        sheet.Cells[i + 1, 2].PutValue(dt[i].ChargeName);
                        sheet.Cells[i + 1, 3].PutValue(dt[i].TotalCount);
                        sheet.Cells[i + 1, 4].PutValue(dt[i].FinishCount);
                        sheet.Cells[i + 1, 5].PutValue(dt[i].NotFinishCount);
                        sheet.Cells[i + 1, 6].PutValue(dt[i].CompletionRate);
                    }
                //    MemoryStream bookStream = new MemoryStream();
                    wb.Save(bookStream, new OoxmlSaveOptions(SaveFormat.Xlsx));
                    bookStream.Seek(0, SeekOrigin.Begin);
                   // return File(bookStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("辖区统计_{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff")));//最后以文件形式返回
                }
                //业管统计 area
                else if (model.QueryType == "kind")
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    string templatePath = string.Format("~\\Excel\\KindCount.xlsx");
                    FileStream fs = new FileStream(System.Web.HttpContext.Current.Server.MapPath(templatePath), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    Workbook wb = new Workbook(fs);
                    Worksheet sheet = wb.Worksheets[0];
                    sheet.Name = "业管统计";
                    fileName = sheet.Name;
                    Cells cells = sheet.Cells;
                    int columnCount = cells.MaxColumn;
                    int rowCount = cells.MaxRow;

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
                        sheet.Cells[i + 1, 0].PutValue(dt[i].Category);
                        sheet.Cells[i + 1, 1].PutValue(dt[i].ChargeName);
                        sheet.Cells[i + 1, 2].PutValue(dt[i].TotalCount);
                        sheet.Cells[i + 1, 3].PutValue(dt[i].FinishCount);
                        sheet.Cells[i + 1, 4].PutValue(dt[i].NotFinishCount);
                        sheet.Cells[i + 1, 5].PutValue(dt[i].CompletionRate);
                    }
                //    MemoryStream bookStream = new MemoryStream();
                    wb.Save(bookStream, new OoxmlSaveOptions(SaveFormat.Xlsx));
                    bookStream.Seek(0, SeekOrigin.Begin);
                 //   return File(bookStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("业管统计_{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff")));//最后以文件形式返回
                }
                   return File(bookStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("{1}_{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff"),fileName));//最后以文件形式返回

            }
            catch (Exception ex)
            {
                return Json(new FlagTips { IsSuccess = false, Msg = ex.Message });
            }
        }
    }
}