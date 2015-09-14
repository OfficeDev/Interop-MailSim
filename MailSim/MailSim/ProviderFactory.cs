using MailSim.Common;
using MailSim.Common.Contracts;
using System;

namespace MailSim
{

    class ProviderFactory
    {
        public static IMailStore CreateMailStore(string mailboxName, MailSimOptions options)
        {
            switch (options.ProviderType)
            {
                case MailSimOptionsProviderType.OOM:
                    return new ProvidersOM.MailStoreProviderOM(mailboxName, options.DisableOutlookPrompts);

                case MailSimOptionsProviderType.HTTP:
                    return new ProvidersREST.MailStoreProviderHTTP(options.UserName, options.Password);

                case MailSimOptionsProviderType.SDK:
                    return new ProvidersREST.MailStoreProviderSDK(options.UserName, options.Password);

                default:
                    throw new Exception(string.Format("Unknown provider type: {0}!", options.ProviderType));
            }
        }
    }
}
