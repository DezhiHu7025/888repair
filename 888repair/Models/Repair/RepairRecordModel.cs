using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models.Repair
{
    public class RepairRecordModel
    {
        /// <summary>
        /// 请修编号
        /// </summary>
        public string RepairId { get; set; }

        /// <summary>
        /// 辖区Id
        /// </summary>
        public string AreaId { get; set; }

        /// <summary>
        /// 业管Id
        /// </summary>
        public string KindId { get; set; }

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
        public string Loaction { get; set; }

        /// <summary>
        /// 负责人员工编号
        /// </summary>
        public string ChargeEmpno { get; set; }

        /// <summary>
        /// 负责人姓名
        /// </summary>
        public string ChargeEmpname { get; set; }

        /// <summary>
        /// 反应人员工编号
        /// </summary>
        public string ResponseEmpno { get; set; }

        /// <summary>
        /// 反应人员工姓名
        /// </summary>
        public string ResponseEmpname { get; set; }

        /// <summary>
        /// 类别
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// 反应内容
        /// </summary>
        public string ResponseContent { get; set; }

        /// <summary>
        /// 回复内容
        /// </summary>
        public string ReplyContent { get; set; }

        /// <summary>
        /// 执行状态
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 填表时间
        /// </summary>
        public DateTime? CreatTime { get; set; }

        /// <summary>
        /// 完成时间
        /// </summary>
        public DateTime? FinishTime { get; set; }

        /// <summary>
        /// 照片位置
        /// </summary>
        public string PhotoPath { get; set; }

        /// <summary>
        /// 空间编号
        /// </summary>
        public string RoomNum { get; set; }

        /// <summary>
        ///维修时段
        /// </summary>
        public string RepairTime { get; set; }

        /// <summary>
        /// 分机号码
        /// </summary>
        public string Telephone { get; set; }

        /// <summary>
        /// 破坏原因
        /// </summary>
        public string DamageReason { get; set; }

        /// <summary>
        /// 破坏人所属单位
        /// </summary>
        public string DamageClass { get; set; }

        /// <summary>
        /// 破坏人姓名
        /// </summary>
        public string DamageName { get; set; }

        /// <summary>
        /// 填表时间开始日期
        /// </summary>
        public DateTime? startDate { get; set; }

        /// <summary>
        /// 填表时间结束日期
        /// </summary>
        public DateTime? endDate { get; set; }


    }
}