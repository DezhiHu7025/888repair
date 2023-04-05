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
        public string Building { get; set; }

        /// <summary>
        /// 位置
        /// </summary>
        public string Location { get; set; }

        public string UpdateUser { get; set; }


        public DateTime? UpdateTime { get; set; }

        public Int32 SortNo { get; set; }
    }
}