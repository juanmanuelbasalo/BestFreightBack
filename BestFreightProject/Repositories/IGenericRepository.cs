﻿using BestFreightProject.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace BestFreightProject.Repositories
{
    public interface IGenericRepository<TEntity> where TEntity : BaseEntity
    {
        IQueryable<TEntity> GetAll();
        TEntity Get(int id);
        void Insert(TEntity entity);
        void Delete(TEntity entity);
        void Update(TEntity entity);
        TEntity Find(Expression<Func<TEntity, bool>> searchTerm);
        IEnumerable<TEntity> FindAll(Expression<Func<TEntity, bool>> searchTerm);
        Task<bool> SaveAsync();
    }
}
