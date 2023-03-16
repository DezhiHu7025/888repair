using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace _888repair.Db
{
    public class RepairDb : DbDapperExtension
    {
        public RepairDb()
             : base(ConfigurationManager.ConnectionStrings["Repair"].ConnectionString)
        {
        }
    }
}