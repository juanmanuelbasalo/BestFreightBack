using MailKit.Security;
using Microsoft.Extensions.Configuration;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Services
{
    public class EmailService : IEmailService
    {
        private readonly string smtpServer;
        private readonly int smtpPort;
        private readonly string fromAddress;
        private readonly string fromAddressTitle;
        private readonly string username;
        private readonly string password;
        private readonly bool enableSsl;
        private readonly bool useDefaultCredentials;

        public EmailService(IConfiguration configuration) 
        {
            smtpServer = configuration["Email:SmtpServer"];
            smtpPort = int.Parse(configuration["Email:SmtpPort"]);
            smtpPort = smtpPort == 0 ? 25 : smtpPort;
            fromAddress = configuration["Email:FromAddress"];
            fromAddressTitle = configuration["FromAddressTitle"];
            username = configuration["Email:SmtpUsername"];
            password = configuration["Email:SmtpPassword"];
            enableSsl = bool.Parse(configuration["Email:EnableSsl"]);
            useDefaultCredentials = bool.Parse(configuration["Email:UseDefaultCredentials"]);
        }

        private string CreateEmailBody()
        {
            var mensajeProveedor = @"You have received this quote request from an importer / exporter user of the bestfreightsearch.com platform
			    Please send your freight quote to the person who is in the attached excel file and not to the mail cotizar@bestfreightsearch.com
				If you need more information, do not hesitate to contact the person in the excel file attached.
				Best Regards,

				*******************************************************************************************

				Usted ha recibido esta solicitud de Cotización de un importador/exportador usuario de la plataforma bestfreightsearch.com
				Por favor, enviar su cotización de flete a la persona que está en el archivo Excel adjunto(solicitud de cotización) y no al correo cotizar@bestfreightsearch.com
				Si necesita más información, no dude en contactar a la persona del archivo Excel adjunto.

                Atentos saludos,
                Management 

                +507 6614-2665 
                consultas@BestFreightSearch.com 
                www.BestFreightSearch.com 
                www.BuscoMejorFlete.com";
            return mensajeProveedor;
        }

        public async Task SendWithAttachmentToClientAsync(string toAddress, FileStream file, bool sendAsync = true)
        {
            var mimeMessage = new MimeMessage(); 
            mimeMessage.From.Add(new MailboxAddress(fromAddressTitle, fromAddress));
            mimeMessage.To.Add(new MailboxAddress(toAddress));
            mimeMessage.Subject = "Solicitud de Cotización de Flete";
            var body = CreateEmailBody();

            var bodyBuilder = new BodyBuilder
            {
                HtmlBody = body
            };
            bodyBuilder.Attachments.Add("Solicitud de Cotización.xls", file);

            mimeMessage.Body = bodyBuilder.ToMessageBody();

            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                client.Connect(smtpServer, smtpPort, SecureSocketOptions.StartTls);

                client.Authenticate(username, password); 
                if (sendAsync)
                {
                    await client.SendAsync(mimeMessage);
                }
                else
                {
                    client.Send(mimeMessage);
                }

                client.Disconnect(true);
            }
        }
        public async Task SendWithAttachmentToProviderAsync(string toAddress, FileStream file, bool sendAsync = true)
        {
            var mimeMessage = new MimeMessage();
            mimeMessage.From.Add(new MailboxAddress(fromAddressTitle, fromAddress));
            mimeMessage.To.Add(new MailboxAddress(toAddress));
            mimeMessage.Subject = "Solicitud de Cotización de Flete";
            var body = CreateEmailBody();

            var bodyBuilder = new BodyBuilder
            {
                HtmlBody = body
            };
            bodyBuilder.Attachments.Add("Solicitud de Cotización", file);

            mimeMessage.Body = bodyBuilder.ToMessageBody();

            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                client.Connect(smtpServer, smtpPort, enableSsl);

                client.Authenticate(username, password);
                if (sendAsync)
                {
                    await client.SendAsync(mimeMessage);
                }
                else
                {
                    client.Send(mimeMessage);
                }

                client.Disconnect(true);
            }
        }
    }
}
