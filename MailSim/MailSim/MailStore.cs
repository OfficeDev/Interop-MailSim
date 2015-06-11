//--------------------------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized
// to use this sample source code. For the terms of the license, please see the
// license agreement between you and Microsoft.
//--------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace MailSim.OL
{
    /// <summary>
    /// Class representing Mailbox Store (i.e. mailbox inside Outlook profile)
    /// </summary>
    public class MailStore
    {
        private Dictionary<string, Outlook.OlDefaultFolders> _folderTypes = new Dictionary<string, Outlook.OlDefaultFolders>();
        private Outlook.Store _store;
        private Outlook.Account _userAccount;
        private Outlook.Application _outlook;

        /// <summary>
        /// Constructor. Initiated from MailConnection object
        /// </summary>
        /// <param name="store">Outlook.Store oibject</param>
         public MailStore(Outlook.Store store)
        {
            _store = store;
            _outlook = store.Application;

            // Initalizes the dictionary of supported default Outlook folders
            _folderTypes.Add("olFolderInbox", Outlook.OlDefaultFolders.olFolderInbox);
            _folderTypes.Add("olFolderDeletedItems", Outlook.OlDefaultFolders.olFolderDeletedItems);
            _folderTypes.Add("olFolderDrafts", Outlook.OlDefaultFolders.olFolderDrafts);
            _folderTypes.Add("olFolderJunk", Outlook.OlDefaultFolders.olFolderJunk);
            _folderTypes.Add("olFolderOutbox", Outlook.OlDefaultFolders.olFolderOutbox);
            _folderTypes.Add("olFolderSentMail", Outlook.OlDefaultFolders.olFolderSentMail);

            // Finding Account associated with this store
            // Loops over the Accounts collection of the current Outlook session.
            _userAccount = null;
            Outlook.Accounts accounts = _outlook.Session.Accounts;

            foreach (Outlook.Account account in accounts)
            {
                // When the e-mail address matches, return the account.
                if (account.SmtpAddress == _store.DisplayName)
                {
                    _userAccount = account;
                }
            }

            if (null == _userAccount)
            {
                throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!", _store.DisplayName));
            }
        }

        /// <summary>
        /// Gets top level (root) folder of the mailbox store
        /// </summary>
        /// <returns>MailFolder object that includes all Folders on top level folder layer</returns>
        public MailFolder GetRootFolder()
        {
            return new MailFolder(_store.GetRootFolder() as Outlook.Folder);
        }

        /// <summary>
        /// Returns one of "Default" folders in the mailbox
        /// For list of possible folders refer to
        /// https://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.outlook.oldefaultfolders(v=vs.15).aspx
        /// This class supports the following: 
        /// "olFolderInbox", "olFolderDeletedItems", "olFolderDrafts", "olFolderJunk", "olFolderOutbox", "olFolderSentMail"
        /// </summary>
        /// <param name="folderName">String representing one of default folders (ex.: "olFolderInbox"). </param>
        /// <returns>MailFolder object or null if string does not match supported value</returns>
        public MailFolder GetDefaultFolder(string folderName)
        {
            Outlook.OlDefaultFolders olFolderType;
            try
            {
                olFolderType = GetFolderTypeByName(folderName);
            }
            catch (KeyNotFoundException)
            {
                return null;
            }

            Outlook.Folder folder = _store.GetDefaultFolder(olFolderType) as Outlook.Folder;
            return new MailFolder(folder);
        }

 
        /// <summary>
        /// Finds the Global Address List associated with the MailStore
        /// </summary>
        /// <returns>OLAddressList for GAL or null if store has no GAL</returns>
        public OLAddressList GetGlobalAddressList()
        {
            string PR_EMSMDB_SECTION_UID = @"http://schemas.microsoft.com/mapi/proptag/0x3D150102";

            if (_store == null)
            {
                throw new ArgumentNullException();
            }

            Outlook.PropertyAccessor oPAStore = _store.PropertyAccessor;
            string storeUID = oPAStore.BinaryToString(oPAStore.GetProperty(PR_EMSMDB_SECTION_UID));

            foreach (Outlook.AddressList addrList in _store.Session.AddressLists)
            {
                Outlook.PropertyAccessor oPAAddrList = addrList.PropertyAccessor;
                string addrListUID = oPAAddrList.BinaryToString(oPAAddrList.GetProperty(PR_EMSMDB_SECTION_UID));

                // Returns addrList if match on storeUID
                // and type is olExchangeGlobalAddressList.
                if (addrListUID == storeUID && addrList.AddressListType ==
                    Outlook.OlAddressListType.olExchangeGlobalAddressList)
                {
                    return new OLAddressList(addrList);
                }
            }

            return null;
        }

        /// <summary>
        /// Creates new MailItem associated with this MailStore. 
        /// Typically used for new mail, to be sent from user account of this MailStore.  
        /// </summary>
        /// <returns>MailSim.OL.MailItem</returns>
        public MailItem NewMailItem()
        {
            MailItem mailItem = new MailItem(_outlook.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem);
            mailItem.OutlookMailItem.SendUsingAccount = _userAccount;
            return mailItem;
        }

        /// <summary>
        /// Display name of the store (mailbox) represented by this object
        /// </summary>
        public string DisplayName
        {
            get
            {
                return _store.DisplayName;
            }
        }

        /// <summary>
        /// Gets the folder type by name
        /// </summary>
        /// <param name="folderName">folder name</param>
        /// <returns>folder</returns>
        private Outlook.OlDefaultFolders GetFolderTypeByName(string folderName)
        {
            return (_folderTypes[folderName]);
        }
    }
}
