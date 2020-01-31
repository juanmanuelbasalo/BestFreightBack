using BestFreightProject.Services;
using Microsoft.Extensions.DependencyInjection;
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
                .AddScoped<IEmailService, EmailService>();
        }
    }
}
