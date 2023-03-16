using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models
{
    public class DirectorModel
    {
        /// <summary>
        /// ID
        /// </summary>
        public string charge_id { get; set; }

        /// <summary>
        /// 工号
        /// </summary>
        public string EmpNo { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string FullName { get; set; }

        /// <summary>
        /// 权限
        /// </summary>
        public string Permissiom { get; set; }
    }
}