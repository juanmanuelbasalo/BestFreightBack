using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Services
{
    public interface IEmailService
    {
        Task SendWithAttachmentToClientAsync(string toAddress, FileStream file, bool sendAsync = true);
        Task SendWithAttachmentToProviderAsync(string toAddress, FileStream file, bool sendAsync = true);
    }
}
