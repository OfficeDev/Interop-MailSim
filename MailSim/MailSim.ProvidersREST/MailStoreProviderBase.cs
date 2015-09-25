using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common;
using MailSim.Common.Contracts;
using Microsoft.Azure.ActiveDirectory.GraphClient;

namespace MailSim.ProvidersREST
{
    public class MailStoreProviderBase
    {
        private readonly ActiveDirectoryClient _adClient;

        private static IDictionary<string, string> _predefinedFolders = new Dictionary<string, string>
        {
            {"olFolderInbox", "Inbox"},
            {"olFolderDeletedItems", "Deleted Items"},
            {"olFolderDrafts", "Drafts"},
            {"olFolderJunk", "Junk Email"},
            {"olFolderOutbox", "Outbox"},
            {"olFolderSentMail", "Sent Items"},
        };

        public MailStoreProviderBase(string userName, string password)
        {
            _adClient = AuthenticationHelper.GetGraphClientAsync(userName, password).GetResult();
        }

        protected IAddressBook GetGAL()
        {
            return new AddressBookProvider(_adClient);
        }

        protected string MapFolderName(string name)
        {
            string folderName;

            if (_predefinedFolders.TryGetValue(name, out folderName) == false)
            {
                return null;
            }

            return folderName;
        }
    }
}
