using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models.Repair
{
    public class StepRecordModel
    {
        /// <summary>
        /// 唯一编号
        /// </summary>
        public string GUID { get; set; }

        /// <summary>
        /// 请修编号
        /// </summary>
        public string RepairId { get; set; }

        /// <summary>
        /// 状态
        /// </summary>
        public string STATUS { get; set; }

        /// <summary>
        /// 意见
        /// </summary>
        public string OPINION { get; set; }

        /// <summary>
        /// 状态
        /// </summary>
        public string STEP { get; set; }

        /// <summary>
        /// 负责人员工编号
        /// </summary>
        public string ChargeEmpno { get; set; }

        /// <summary>
        /// 负责人姓名
        /// </summary>
        public string ChargeEmpname { get; set; }
        
        /// <summary>
        /// 记录人工号
        /// </summary>
        public string UpdateEmpNo { get; set; }

        /// <summary>
        /// 记录人姓名
        /// </summary>
        public string UpdateEmpName { get; set; }

        /// <summary>
        /// 记录时间
        /// </summary>
        public DateTime? UpdateTime { get; set; }

        /// <summary>
        /// 自增标志 排序用
        /// </summary>
        public string SORT { get; set; }
    }
}