using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models.State
{
    public class StateModel
    {
        /// <summary>
        /// id
        /// </summary>
        public string ID { get; set; }

        /// <summary>
        /// 狀態
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 權限
        /// </summary>
        public string Permissiom { get; set; }
    }
}