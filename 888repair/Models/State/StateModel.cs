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

        public string SystemCategory { get; set; }


        public string StatusValue { get; set; }

        public string StatusText { get; set; }

        public string UpdateUser { get; set; }


        public DateTime? UpdateTime { get; set; }
    }
}