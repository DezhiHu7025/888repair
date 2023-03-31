using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models.Kind
{
    public class KindMatchModel
    {

        public string MatchType { get; set; }

        public string MatchId { get; set; }

        public string AreaId { get; set; }

        public string ChargeId { get; set; }

        public string Sort { get; set; }

        public string UpdateUser { get; set; }

        public DateTime? UpdateTime { get; set; }

        public string EmpNo { get; set; }

        public string FullName { get; set; }

        public string KindCategory { get; set; }

        public string Remark { get; set; }

        public Int32 SortNo { get; set; }

        public string SystemCategory { get; set; }
    }
}