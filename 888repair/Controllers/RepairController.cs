using System;
using System.Collections.Generic;
using System.Linq;
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
    }
}