using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models.Director
{
    public class UserModel
    {
        public string EmpNo { get; set; }

        public string Account { get; set; }

        public string Ename { get; set; }

        public string Cname { get; set; }

        public string fullname { get; set; }

        public string titlename { get; set; }

        public string status { get; set; }

        public string deptid2 { get; set; }

        public string DeptName { get; set; }

        public string sourcetype { get; set; }

        public string email { get; set; }

        public string Password { get; set; }

        public string GroupId { get; set; }

        /// <summary>
        /// 资讯or后勤权限
        /// </summary>
        public string GroupName { get; set; }

        public string password2 { get; set; }

        /// <summary>
        /// 浏览者权限
        /// </summary>
        public string Viewer { get; set; }

        /// <summary>
        /// 管理者权限
        /// </summary>
        public string Manager { get; set; }

    }
}