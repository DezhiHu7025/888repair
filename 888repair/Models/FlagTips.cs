using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace _888repair.Models
{
    public class FlagTips
    {
        public object data { get; set; }

        public bool IsSuccess { get; set; }
        public string Msg { get; set; }

        public int code { get; set; }

        public int count { get; set; }
        public string FileName { get; set; }
        public string FileType { get; set; }
    }
}