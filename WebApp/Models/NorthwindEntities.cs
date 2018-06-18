using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace WebApp.Models
{
    public class NorthwindContext : DbContext
    {
        public NorthwindContext() : base("NorthwindEntities")
        {

        }
        public DbSet<Product> Products { get; set; }
    }
}