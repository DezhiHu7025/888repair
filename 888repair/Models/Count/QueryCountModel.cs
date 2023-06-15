using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models.Count
{
    public class QueryCountModel
    {
        public string SystemCategory { get; set; }

        public string QueryType { get; set; }

        public DateTime? StartDate { get; set; }

        public DateTime? EndDate { get; set; }
    }
}