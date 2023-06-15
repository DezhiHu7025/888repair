using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models.Count
{
    public class CountModel
    {
        /// <summary>
        /// 总数
        /// </summary>
        public string TotalCount { get; set; }

        /// <summary>
        /// 完成数量
        /// </summary>
        public string FinishCount { get; set; }

        /// <summary>
        /// 未完成数量
        /// </summary>
        public string NotFinishCount { get; set; }

        /// <summary>
        /// 执行率
        /// </summary>
        public string CompletionRate { get; set; }

        /// <summary>
        /// 负责人姓名
        /// </summary>
        public string ChargeName { get; set; }

        /// <summary>
        /// 反应人姓名
        /// </summary>
        public string ResponseName { get; set; }

        /// <summary>
        /// 辖区：大楼别
        /// </summary>
        public string Building { get; set; }

        /// <summary>
        /// 辖区：位置
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// 业管：类别
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// 反应人部门
        /// </summary>
        public string Department { get; set; }
    }
}