using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace _888repair.Db
{
    public class EmailDb : DbDapperExtension
    {
        public EmailDb()
             : base(ConfigurationManager.ConnectionStrings["Email"].ConnectionString)
        {
        }
    }
}