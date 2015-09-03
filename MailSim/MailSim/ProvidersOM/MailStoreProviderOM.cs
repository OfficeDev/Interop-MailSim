﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

using Outlook = Microsoft.Office.Interop.Outlook;
using MailSim.Contracts;
using System.Diagnostics;
using Microsoft.Win32;

namespace MailSim.ProvidersOM
{
    class MailStoreProviderOM : IMailStore
    {
        private const string OfficeVersion = "15.0";
        private const string OfficePolicyRegistryRoot = @"Software\Policies\Microsoft\Office\" + OfficeVersion;
        private const string OutlookPolicyRegistryRoot = OfficePolicyRegistryRoot + @"\Outlook";

        private readonly Outlook.Store _store;
        private readonly Outlook.Account _userAccount;
        private Outlook.Application _outlook;
        private bool _keepOutlookRunning = false;
        private readonly bool _disableOutlookPrompt;
        private readonly Lazy<IMailFolder> _rootFolder;

        private Dictionary<string, Outlook.OlDefaultFolders> _folderTypes = new Dictionary<string, Outlook.OlDefaultFolders>
        {
            {"olFolderInbox", Outlook.OlDefaultFolders.olFolderInbox},
            {"olFolderDeletedItems", Outlook.OlDefaultFolders.olFolderDeletedItems},
            {"olFolderDrafts", Outlook.OlDefaultFolders.olFolderDrafts},
            {"olFolderJunk", Outlook.OlDefaultFolders.olFolderJunk},
            {"olFolderOutbox", Outlook.OlDefaultFolders.olFolderOutbox},
            {"olFolderSentMail", Outlook.OlDefaultFolders.olFolderSentMail},
        };

        internal MailStoreProviderOM(string mailboxName, bool disableOutlookPrompt)
        {
            _disableOutlookPrompt = disableOutlookPrompt;

            if (_disableOutlookPrompt)
            {
                ConfigOutlookPrompts(false);
            }

            ConnectToOutlook();

            if (mailboxName != null)
            {
                mailboxName = mailboxName.ToLower();
                _store = AllMailStores().FirstOrDefault(x => x.DisplayName.ToLower() == mailboxName);

                if (_store == null)
                {
                    throw new ArgumentException(string.Format("Cannot find store (mailbox) {0} in default profile", mailboxName));
                }
            }
            else
            {
                _store = _outlook.Session.DefaultStore;
            }

            _userAccount = FindUserAccount();
            _rootFolder = new Lazy<IMailFolder>(GetRootFolder);
        }

        public string DisplayName
        {
            get
            {
                return _store.DisplayName;
            }
        }

        private Outlook.Account FindUserAccount()
        {
            Outlook.Accounts accounts = _outlook.Session.Accounts;

            foreach (Outlook.Account account in accounts)
            {
                if (account.SmtpAddress == _store.DisplayName)
                {
                    return account;
                }
            }

            throw new System.Exception(string.Format("No Account with SmtpAddress: {0} exists!", _store.DisplayName));
        }

        private IEnumerable<Outlook.Store> AllMailStores()
        {
            // Get all mailboxes (stores) in the profile. 
            //Returns only email stores (skips Public Folders, Delegates, Archives, PSTs)
            foreach (Outlook.Store store in _outlook.Session.Stores)
            {
                if ((store.ExchangeStoreType == Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox)
                    || (store.ExchangeStoreType == Outlook.OlExchangeStoreType.olAdditionalExchangeMailbox)
                    || (store.ExchangeStoreType == Outlook.OlExchangeStoreType.olNotExchange))
                {
                    yield return store;
                }
            }
        }

        public IMailFolder GetDefaultFolder(string folderName)
        {
            Outlook.OlDefaultFolders olFolderType;

            if (_folderTypes.TryGetValue(folderName, out olFolderType) == false)
            {
                return null;
            }

            Outlook.Folder folder = _store.GetDefaultFolder(olFolderType) as Outlook.Folder;
            return new MailFolderProviderOM(folder);
        }

        public IMailFolder RootFolder
        {
            get
            {
                return _rootFolder.Value;
            }
        }

        private IMailFolder GetRootFolder()
        {
            return new MailFolderProviderOM(_store.GetRootFolder() as Outlook.Folder);
        }

        public IMailItem NewMailItem()
        {
            var mailItem = new MailItemProviderOM(_outlook.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem);
            mailItem.Handle.SendUsingAccount = _userAccount;
            return mailItem;
        }

        private void ConnectToOutlook()
        {
            // Checks whether an Outlook process is currently running
            try
            {
                if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
                {
                    Log.Out(Log.Severity.Info, "Connection", "Connecting to an existing Outlook instance");
                    _outlook = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                    _keepOutlookRunning = true;
                    return;
                }

                // Creates a new instance of Outlook and logs on to the specified profile.
                Log.Out(Log.Severity.Info, "Connection", "Starting a new Outlook session");

                _outlook = new Outlook.Application();
                Outlook.NameSpace nameSpace = _outlook.GetNamespace("MAPI");
                Outlook.Folder mailFolder = (Outlook.Folder)nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            }
            catch (Exception)
            {
                Log.Out(Log.Severity.Error, "Connection", "Error encountered when connecting to Outlook ");
                throw;
            }
        }

        /// <summary>
        /// Finds the Global Address List associated with the MailStore
        /// </summary>
        /// <returns>OLAddressList for GAL or null if store has no GAL</returns>
        public IAddressBook GetGlobalAddressList()
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
                    return new AddressBookProviderOM(addrList);
                }
            }

            return null;
        }

        ~MailStoreProviderOM()
        {
            // Restore the Outlook prompt if needed 
            if (_disableOutlookPrompt == true)
            {
                ConfigOutlookPrompts(true);
            }

            // Closes the Outlook process
            if (_outlook != null && !_keepOutlookRunning)
            {
                Console.WriteLine("Exiting Outlook");
 
                ((Outlook._Application)_outlook).Quit();
            }
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

        /// <summary>
        /// This method updates the registry to turn on/off Outlook prompts.
        /// This is documented in http://support.microsoft.com/kb/926512
        /// </summary>
        /// <param name="show">True to enable Outlook prompts, False to disable Outlook prompts.</param>
        private static void ConfigOutlookPrompts(bool show)
        {
            const string olSecurityKey = @"HKEY_CURRENT_USER\" + OutlookPolicyRegistryRoot + @"\Security";
            const string adminSecurityMode = "AdminSecurityMode";
            const string addressBookAccess = "PromptOOMAddressBookAccess";
            const string addressInformationAccess = "PromptOOMAddressInformationAccess";
            const string saveAs = "PromptOOMSaveAs";
            const string customAction = "PromptOOMCustomAction";
            const string send = "PromptOOMSend";
            const string meetingRequestResponse = "PromptOOMMeetingTaskRequestResponse";

            try
            {
                if (show == true)
                {
                    Registry.SetValue(olSecurityKey, adminSecurityMode, (int)0);
                    Registry.SetValue(olSecurityKey, addressBookAccess, (int)1);
                    Registry.SetValue(olSecurityKey, addressInformationAccess, (int)1);
                    Registry.SetValue(olSecurityKey, customAction, (int)1);
                    Registry.SetValue(olSecurityKey, saveAs, (int)1);
                    Registry.SetValue(olSecurityKey, send, (int)1);
                    Registry.SetValue(olSecurityKey, meetingRequestResponse, (int)1);
                }
                else
                {
                    Registry.SetValue(olSecurityKey, adminSecurityMode, (int)3);
                    Registry.SetValue(olSecurityKey, addressBookAccess, (int)2);
                    Registry.SetValue(olSecurityKey, addressInformationAccess, (int)2);
                    Registry.SetValue(olSecurityKey, customAction, (int)2);
                    Registry.SetValue(olSecurityKey, saveAs, (int)2);
                    Registry.SetValue(olSecurityKey, send, (int)2);
                    Registry.SetValue(olSecurityKey, meetingRequestResponse, (int)2);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, "", "Unable to change registry, you may want to run this as Administrator\n" + ex.ToString());
            }
        }
    }
}
