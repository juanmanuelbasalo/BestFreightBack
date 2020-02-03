using BestFreightProject.Entities;
using BestFreightProject.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace BestFreightProject.Services
{
    public class FreightProviderService : IFreightProviderService
    {
        private readonly IGenericRepository<FreightProvider> repository;
        public FreightProviderService(IGenericRepository<FreightProvider> repository) => this.repository = repository; 
        public Task<bool> DeleteProviders(FreightProvider entity)
        {
            throw new NotImplementedException();
        }

        public FreightProvider FindProviders(Expression<Func<FreightProvider, bool>> searchTerm)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<FreightProvider> GetAllProviders()
        {
            var provider = repository.GetAll().Where(provider => 
                                (!string.IsNullOrEmpty(provider.Email) || !string.IsNullOrEmpty(provider.Email)) && provider.Status == 1).ToList();
            return provider;
        }

        public FreightProvider GetProviders(Guid id)
        {
            throw new NotImplementedException();
        }

        public Task<FreightProvider> InsertProviders(FreightProvider entity)
        {
            throw new NotImplementedException();
        }

        public Task<FreightProvider> UpdateProvidersAsync(FreightProvider ProvidersDto)
        {
            throw new NotImplementedException();
        }
    }
}
