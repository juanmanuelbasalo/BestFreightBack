using BestFreightProject.Entities;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Database
{
    public class BestFreightContext : DbContext
    {
        public BestFreightContext(DbContextOptions<BestFreightContext> options) : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
            modelBuilder.Entity<FreightProvider>().HasKey(entity => entity.Id);
        }

    }
}
