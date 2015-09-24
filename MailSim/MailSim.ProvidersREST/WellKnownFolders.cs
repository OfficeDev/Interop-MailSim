using System.Collections.Generic;

namespace MailSim.ProvidersREST
{
    class WellKnownFolders
    {
        private static IDictionary<string, string> _predefinedFolders = new Dictionary<string, string>
        {
            {"olFolderInbox", "Inbox"},
            {"olFolderDeletedItems", "Deleted Items"},
            {"olFolderDrafts", "Drafts"},
            {"olFolderJunk", "Junk Email"},
            {"olFolderOutbox", "Outbox"},
            {"olFolderSentMail", "Sent Items"},
        };

        internal static string MapFolderName(string name)
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
