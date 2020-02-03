using BestFreightProject.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace BestFreightProject.Services
{
    public interface IFreightProviderService
    {
        IEnumerable<FreightProvider> GetAllProviders();
        FreightProvider GetProviders(Guid id);
        Task<FreightProvider> InsertProviders(FreightProvider entity);
        Task<FreightProvider> UpdateProvidersAsync(FreightProvider userDto);
        Task<bool> DeleteProviders(FreightProvider entity);
        FreightProvider FindProviders(Expression<Func<FreightProvider, bool>> searchTerm);
    }
}
