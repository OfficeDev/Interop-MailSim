using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSim.Contracts
{
    interface IMailStore
    {
        /// <summary>
        /// Creates new MailItem associated with this MailStore. 
        /// Typically used for new mail, to be sent from user account of this MailStore.  
        /// </summary>
        /// <returns>IMailItem</returns>
        IMailItem NewMailItem();
        /// <summary>
        /// Display name of the store (mailbox) represented by this object
        /// </summary>
        string DisplayName { get; }
        /// <summary>
        /// Top level (root) folder of the mailbox store
        /// </summary>
        /// <returns>IMailFolder object that includes all Folders on top level folder layer</returns>
        IMailFolder RootFolder { get; }
        /// <summary>
        /// Returns one of "Default" folders in the mailbox
        /// For list of possible folders refer to
        /// https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.outlook.oldefaultfolders(v=vs.15).aspx
        /// This class supports the following: 
        /// "olFolderInbox", "olFolderDeletedItems", "olFolderDrafts", "olFolderJunk", "olFolderOutbox", "olFolderSentMail"
        /// </summary>
        /// <param name="folderName">String representing one of default folders (ex.: "olFolderInbox"). </param>
        /// <returns>MailFolder object or null if string does not match supported value</returns>
        IMailFolder GetDefaultFolder(string name);
        /// <summary>
        /// Finds the Global Address List associated with the MailStore
        /// </summary>
        /// <returns>IAddressBook for GAL or null if store has no GAL</returns>
        IAddressBook GetGlobalAddressList();
    }
}
