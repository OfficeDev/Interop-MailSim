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
            var user = HttpUtil.GetItemAsync<User>(string.Empty).Result;
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
            var newMessage = HttpUtil.PostItemAsync("Messages", message).Result;

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
            return GetGAL();
        }

        private class User
        {
            public string Id { get; set; }
        }
    }
}
