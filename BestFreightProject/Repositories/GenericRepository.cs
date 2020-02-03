﻿using BestFreightProject.Database;
using BestFreightProject.Entities;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace BestFreightProject.Repositories
{
    public class GenericRepository<TEntity> : IGenericRepository<TEntity> where TEntity : BaseEntity
    {
        private readonly BestFreightContext context;
        private DbSet<TEntity> entities;
        public GenericRepository(BestFreightContext context) => this.context = context;

        private DbSet<TEntity> Entities => entities ?? (entities = context.Set<TEntity>());

        public void Delete(TEntity entity) => Entities.Remove(entity);
        public TEntity Get(int id) => Entities.AsNoTracking().FirstOrDefault(entity => entity.Id == id);
        public IQueryable<TEntity> GetAll() => Entities.AsNoTracking();
        public void Insert(TEntity entity) => Entities.Add(entity);
        public TEntity Find(Expression<Func<TEntity, bool>> searchTerm) => Entities.SingleOrDefault(searchTerm);
        public void Update(TEntity entity) => context.Update(entity);
        public async Task<bool> SaveAsync()
        {
            var result = await context.SaveChangesAsync();
            return result >= 0;
        }
        public IEnumerable<TEntity> FindAll(Expression<Func<TEntity, bool>> searchTerm) => Entities.Where(searchTerm);
    }
}

