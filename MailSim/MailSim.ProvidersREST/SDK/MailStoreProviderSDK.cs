using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common;
using MailSim.Common.Contracts;
using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.Discovery;

namespace MailSim.ProvidersREST
{
    public class MailStoreProviderSDK : MailStoreProviderBase, IMailStore
    {
        private readonly Uri DiscoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");
        private const string DiscoveryResourceId = "https://api.office.com/discovery/";

        private OutlookServicesClient _outlookClient;
        private readonly IUser _user;

        public MailStoreProviderSDK(string userName, string password) :
            base(userName, password)
        {
            _outlookClient = GetOutlookClient("Mail");

            _user = _outlookClient.Me.ExecuteAsync().Result;

            DisplayName = _user.Id;
            RootFolder = new MailFolderProviderSDK(_outlookClient, _user.Id);
        }

        public string DisplayName { get; private set; }

        public IMailFolder RootFolder { get; private set; }

        public IMailItem NewMailItem()
        {
            ItemBody body = new ItemBody
            {
                Content = "New Body",
                ContentType = BodyType.HTML
            };

            Message message = new Message
            {
                Subject = "New Subject",
                Body = body,
                ToRecipients = new List<Recipient>(),
                Importance = Importance.High
            };

            // Save the draft message. Saving to Me.Messages saves the message in the Drafts folder.
            _outlookClient.Me.Messages.AddMessageAsync(message).Wait();

            return new MailItemProviderSDK(_outlookClient, message);
        }

        public IMailFolder GetDefaultFolder(string name)
        {
            string folderName = MapFolderName(name);

            if (folderName == null)
            {
                return null;
            }

            return RootFolder.SubFolders.FirstOrDefault(x => x.Name == folderName);
        }

        /// <summary>
        /// Finds the Global Address List associated with the MailStore
        /// </summary>
        /// <param name="folderPath">parameter for FindFolder</param>
        /// <returns>IMailFolder for the target of the event</returns>
        /// In this case the Folder element is a path in Outlook. Each component of the path separated by '\'.
        /// The first or default folder in the path, can be preceded by "\\" or nothing. If it's the only part of
        /// the path, then it MUST be one of the default OL folders (see the schema for the EventMonitor operation). 
        public IMailFolder FindFolder(string folderPath)
        {
            IMailFolder folder;
            System.Char backslash = '\\';
            if (folderPath.StartsWith(@"\\"))
            {
                folderPath = folderPath.Remove(0, 2);
            }
            String[] folders = folderPath.Split(backslash);
            folder = GetDefaultFolder(folders[0]);

            if (folder != null)
            {
                for (int i = 1; i <= folders.Length - 1; i++)
                {
                    IEnumerable<IMailFolder> subFolders = folder.SubFolders;
                    folder = subFolders.FirstOrDefault(fld => fld.Name.Equals(folders[i], StringComparison.CurrentCultureIgnoreCase));
                    if (folder == null)
                    {
                        return null;
                    }
                }
            }
            return folder;
        }

        public IAddressBook GetGlobalAddressList()
        {
            return base.GetGAL();
        }

        private OutlookServicesClient GetOutlookClient(string capability)
        {
            if (_outlookClient != null)
            {
                return _outlookClient;
            }

            try
            {
                Uri serviceEndpointUri;
                string serviceResourceId;

                GetService(capability, out serviceEndpointUri, out serviceResourceId);

                _outlookClient = new OutlookServicesClient(
                    serviceEndpointUri,
                    async () => await AuthenticationHelper.GetTokenAsync(serviceResourceId));
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Warning, string.Empty, ex.ToString());
            }

            return _outlookClient;
        }

        private void GetService(string capability, out Uri serviceEndpointUri, out string serviceResourceId)
        {
            var discoveryClient = new DiscoveryClient(DiscoveryServiceEndpointUri,
                async () => await AuthenticationHelper.GetTokenAsync(DiscoveryResourceId));

            CapabilityDiscoveryResult result = discoveryClient.DiscoverCapabilityAsync(capability).Result;
            serviceEndpointUri = result.ServiceEndpointUri;
            serviceResourceId = result.ServiceResourceId;
        }
    }
}
