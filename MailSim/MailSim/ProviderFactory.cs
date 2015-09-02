using MailSim.Common.Contracts;

namespace MailSim
{
    // TODO: Define the provider selection via config file
    class ProviderFactory
    {
        public static IMailStore CreateMailStore(string mailboxName, MailSimSequence seq = null)
        {
            if (true)
            {
                return new ProvidersOM.MailStoreProviderOM(mailboxName, seq == null ? false : seq.DisableOutlookPrompt);
            }
            else
            {
//                return new ProvidersREST.MailStoreProviderSDK();
                return new ProvidersREST.MailStoreProviderHTTP();
            }
        }
    }
}
