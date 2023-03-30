using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models.Kind
{
    public class KindModel
    {
        public string KindID { get; set; }

        /// <summary>
        /// 系统类别
        /// </summary>
        public string SystemCategory { get; set; }

        /// <summary>
        /// 类别
        /// </summary>
        public string KindCategory { get; set; }

        /// <summary>
        /// 备注
        /// </summary>
        public string Remark { get; set; }
        
        /// <summary>
        /// 排序
        /// </summary>
        public string Sort { get; set; }

        public string UpdateUser { get; set; }


        public DateTime? UpdateTime { get; set; }
    }
}