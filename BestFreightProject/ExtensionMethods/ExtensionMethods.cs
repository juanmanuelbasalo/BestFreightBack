using BestFreightProject.Database;
using BestFreightProject.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Pomelo.EntityFrameworkCore.MySql.Infrastructure;
using Pomelo.EntityFrameworkCore.MySql.Storage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.ExtensionMethods
{
    public static class ExtensionMethods
    {
        public static IServiceCollection AddScopedCustomServices(this IServiceCollection service)
        {
            return service.AddScoped<IExcelService, ExcelService>()
                            .AddScoped<IEmailService, EmailService>()
                                .AddScoped<IFreightProviderService, FreightProviderService>();
        }
        public static IServiceCollection AddCustomDbContext(this IServiceCollection service, IConfiguration configuration)
        {
            return service.AddDbContextPool<BestFreightContext>( options =>
            {
                options.UseLazyLoadingProxies();
                options.UseMySql(configuration.GetConnectionString("DefaultConnection"));
                options.EnableSensitiveDataLogging();
            });
        }
    }
}
