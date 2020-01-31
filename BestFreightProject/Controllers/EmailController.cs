using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using BestFreightProject.Dtos;
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

        private readonly ILogger<EmailController> _logger;

        public EmailController(ILogger<EmailController> logger, IExcelService excelService, IEmailService emailService)
        {
            _logger = logger;
            this.excelService = excelService;
            this.emailService = emailService;
        }

        [HttpPost("create")]
        [ProducesResponseType(201)]
        [ProducesResponseType(500)]
        [ProducesResponseType(400)]
        [ProducesResponseType(401)]
        public async Task<ActionResult> SendEmail([FromBody] ExcelCreateDto excelCreateDto)
        {
            var succesfulCreate = excelService.CreateExcelFile(excelCreateDto);
            if (succesfulCreate)
            {
                var file = excelService.GetExcel(@"c:\temp\MyTest.xls");
                await emailService.SendWithAttachmentToClientAsync(@"jb_2896@hotmail.com",file);
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
    }
}
