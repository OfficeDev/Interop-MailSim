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
using System.Linq;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailSim.OL
{
    public class MailConnection
    {
        #region Private Variables
        private Outlook._Application _outlook;
        private bool _keepOutlookRunning = false;
        #endregion

        /// <summary>
        /// Constructor
        /// </summary>
        public MailConnection()
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
        /// Destructor
        /// </summary>
        ~MailConnection()
        {
            // Closes the Outlook process
            if (_outlook != null && !_keepOutlookRunning)
            {
                Console.WriteLine("Exiting Outlook");
 
                _outlook.Quit();
            }
        }

        /// <summary>
        /// This method gets all mail stores
        /// </summary>
        /// <returns>an enumeration of mail store objects</returns>
        public IEnumerable<MailStore> GetAllMailStores()
        {
            foreach (Outlook.Store store in _outlook.Session.Stores)
            {
                if ((store.ExchangeStoreType == Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox)
                    || (store.ExchangeStoreType == Outlook.OlExchangeStoreType.olAdditionalExchangeMailbox)
                    || (store.ExchangeStoreType == Outlook.OlExchangeStoreType.olNotExchange))
                {
                    yield return new MailStore(store);
                }
            }
        }

        /// <summary>
        /// This method gets the default mail store
        /// </summary>
        /// <returns>default mail store</returns>
        public MailStore GetDefaultStore()
        {
            return new MailStore(_outlook.Session.DefaultStore);        
        }
    }
}
