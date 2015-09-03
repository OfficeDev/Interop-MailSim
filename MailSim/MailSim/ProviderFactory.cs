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
        public static IMailStore CreateMailStore(string mailboxName, MailSimSequence seq=null)
        {
            // Opens connection to Outlook with default profile, starts Outlook if it is not running
            // Note: Currently only the default MailStore is supported.
            return new MailStoreProviderOM(mailboxName, seq == null ? false : seq.DisableOutlookPrompt);
        }
    }
}
