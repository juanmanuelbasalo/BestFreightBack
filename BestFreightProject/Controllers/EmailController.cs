using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BestFreightProject.Dtos;
using BestFreightProject.Entities;
using BestFreightProject.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace BestFreightProject.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class EmailController : ControllerBase
    {
        private readonly IExcelService excelService;
        private readonly IEmailService emailService;
        private readonly IFreightProviderService freightProviderService;

        private readonly ILogger<EmailController> _logger;

        public EmailController(ILogger<EmailController> logger, IExcelService excelService, 
            IEmailService emailService, IFreightProviderService freightProviderService)
        {
            _logger = logger;
            this.excelService = excelService;
            this.emailService = emailService;
            this.freightProviderService = freightProviderService;
        }

        [HttpPost("create")]
        [ProducesResponseType(200)]
        [ProducesResponseType(500)]
        [ProducesResponseType(400)]
        [ProducesResponseType(401)]
        public async Task<ActionResult> SendEmailAsync([FromBody] ExcelCreateDto excelCreateDto)
        {
            var succesfulCreate = excelService.CreateExcelFile(excelCreateDto);
            if (succesfulCreate)
            {
                var providers = freightProviderService.GetAllProviders().Where(pro => 
                                  pro.Country.Equals(excelCreateDto.CompanyInfo.Country) 
                                    && pro.Service.Equals(excelCreateDto.QuotationInfo.SubFreightType)).ToList();

                var file = excelService.GetExcel(@"c:\temp\MyTest.xls");

                foreach (var provider in providers)
                {
                    var email = provider.Email.Equals("") ? provider.Email2 : provider.Email; 
                    await emailService.SendWithAttachmentToClientAsync(@$"{email}", file);
                }

                return Ok();
            }

            throw new Exception("Something went wrong creating the excel.");
        }

        [HttpGet]
        public ActionResult<ExcelCreateDto> GetExcelAsJsonFormat()
        {
            var excelDto = new ExcelCreateDto
            {
                CargoDeliveryInfo = new Entities.CargodeliveryInformation(),
                CargoInfo = new Entities.CargoInformation(),
                CargoReceiptInfo = new Entities.CargoreceiptInformation(),
                CompanyInfo = new Entities.CompanyInformation(),
                OceanCarriersInfo = new Entities.OceanCarriers { LogisticServices = new List<Entities.LogisticService> { new Entities.LogisticService { Name = "Test" } } },
                QuotationInfo = new Entities.QuotationInformation()
            };

            var jsonExcel = JsonConvert.SerializeObject(excelDto);
            return Ok(jsonExcel);
        }

        [HttpGet("user")]
        public ActionResult<List<FreightProvider>> GetAllUsers()
        {
            return freightProviderService.GetAllProviders().ToList();
        }
    }
}
