using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common.Contracts;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Newtonsoft.Json.Linq;

namespace MailSim.ProvidersREST
{
    public class MailStoreProviderHTTP : MailStoreProviderBase, IMailStore
    {
        public MailStoreProviderHTTP(string userName, string password) :
            base(userName, password)
        {
            var user = HttpUtilSync.GetItem<User>(string.Empty);
            DisplayName = user.Id;
            RootFolder = new MailFolderProviderHTTP(null, DisplayName);
        }

        public string DisplayName { get; private set; }

        public IMailFolder RootFolder { get; private set; }

        public IMailItem NewMailItem()
        {
            var body = new MailItemProviderHTTP.ItemBody
            {
                Content = "New Body",
                ContentType = "HTML"
            };

            var message = new MailItemProviderHTTP.Message
            {
                Subject = "New Subject",
                Body = body,
                ToRecipients = new List<MailItemProviderHTTP.Recipient>(),
                Importance = "High"
            };

            // Save the draft message.
            var newMessage = HttpUtilSync.PostItem("Messages", message);

            return new MailItemProviderHTTP(newMessage);
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

        public IAddressBook GetGlobalAddressList()
        {
            return GetGAL();
        }

        private class User
        {
            public string Id { get; set; }
        }
    }
}
