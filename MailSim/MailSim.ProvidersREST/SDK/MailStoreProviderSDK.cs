using System;
using System.Collections.Generic;
using System.Linq;
using MailSim.Common;
using MailSim.Common.Contracts;
using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.Discovery;
using Microsoft.Azure.ActiveDirectory.GraphClient;

namespace MailSim.ProvidersREST
{
    public class MailStoreProviderSDK : IMailStore
    {
        private readonly Uri DiscoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");
        private const string DiscoveryResourceId = "https://api.office.com/discovery/";
        private readonly ActiveDirectoryClient _adClient;

        private OutlookServicesClient _outlookClient;
        private readonly Microsoft.Office365.OutlookServices.IUser _user;

        public MailStoreProviderSDK(string userName, string password)
        {
            _adClient = AuthenticationHelperSDK.GetGraphClientAsync(userName, password).GetResult();
            _outlookClient = GetOutlookClient("Mail");

            _user = _outlookClient.Me.ExecuteAsync().GetResult();

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
            _outlookClient.Me.Messages.AddMessageAsync(message).GetResult();

            return new MailItemProviderSDK(_outlookClient, message);
        }

        public IMailFolder GetDefaultFolder(string name)
        {
            string folderName = WellKnownFolders.MapFolderName(name);

            if (folderName == null)
            {
                return null;
            }

            return RootFolder.SubFolders.FirstOrDefault(x => x.Name == folderName);
        }

        public IAddressBook GetGlobalAddressList()
        {
            return new AddressBookProviderSDK(_adClient);
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
                    async () => await AuthenticationHelperSDK.GetTokenAsync(serviceResourceId));
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
                async () => await AuthenticationHelperSDK.GetTokenAsync(DiscoveryResourceId));

            CapabilityDiscoveryResult result = discoveryClient.DiscoverCapabilityAsync(capability).Result;
            serviceEndpointUri = result.ServiceEndpointUri;
            serviceResourceId = result.ServiceResourceId;
        }
    }
}
