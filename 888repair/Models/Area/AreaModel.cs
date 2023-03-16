using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models.Area
{
    public class AreaModel
    {
        /// <summary>
        /// 辖区ID
        /// </summary>
        public string AreaId { get; set; }

        /// <summary>
        /// 系统类别
        /// </summary>
        public string SystemCategory { get; set; }

        /// <summary>
        /// 大楼别
        /// </summary>
        public string Buliding { get; set; }

        /// <summary>
        /// 位置
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// 权限
        /// </summary>
        public string Permission { get; set; }

        /// <summary>
        /// OU
        /// </summary>
        public string OU { get; set; }
    }
}