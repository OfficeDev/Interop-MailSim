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
using System.Collections;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace MailSim.OL
{
    /// <summary>
    /// Class representing collection of MailFolder ojects
    /// </summary>
    public class MailFolders
    {
        private Outlook.Folders _folders;

        /// <summary>
        /// Constructor. Initiated from MailFolder objects
        /// </summary>
        /// <param name="folders">Outlook.Folders object</param>
        public MailFolders(Outlook.Folders folders)
        {
            _folders = folders;
        }

        /// <summary>
        /// Enumerator. Allows "foreach (MailFolder folder in MailFolders folders) {}"
        /// </summary>
        /// <returns>MailFolder object</returns>
        public IEnumerator GetEnumerator()
        {
            foreach (Outlook.Folder folder in _folders)
            {
                yield return new MailFolder(folder);
            }
        }

        /// <summary>
        /// Gets the first folder
        /// </summary>
        /// <returns>first folder</returns>
        public MailFolder GetFirst()
        {
            return (new MailFolder(_folders.GetFirst() as Outlook.Folder));
        }

        /// <summary>
        /// Gets the last folder
        /// </summary>
        /// <returns>last folder</returns>
        public MailFolder GetLast()
        {
            return (new MailFolder(_folders.GetLast() as Outlook.Folder));
        }

        /// <summary>
        /// Gets the next folder
        /// </summary>
        /// <returns>next folder</returns>
        public MailFolder GetNext()
        {
            return (new MailFolder(_folders.GetNext() as Outlook.Folder));
        }

        /// <summary>
        /// Gets the previous folder
        /// </summary>
        /// <returns>previous folder</returns>
        public MailFolder GetPrevious()
        {
            return (new MailFolder(_folders.GetPrevious() as Outlook.Folder));
        }

        /// <summary>
        /// Count of MailFolder objects in the collection
        /// </summary>
        public int Count
        {
            get 
            { 
                return _folders.Count; 
            } 
        }
    }
}
