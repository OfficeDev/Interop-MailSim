using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MailSim.Contracts;
using MailSim.ProvidersOM;

namespace MailSim
{
    class ProviderFactory
    {
        public static IMailStore CreateMailStore(string mailboxName)
        {
            return new MailStoreProviderOM(mailboxName);
        }
    }
}
