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
    /// Class representing mailbox folder
    /// </summary>
    public class MailFolder
    {
        /// Event Handler: ItemAdd
        public delegate void FolderItemAddEvent(MailItem mailitem);
        private FolderItemAddEvent _itemAddCallback = null;

        private Outlook.Folder _folder;

        /// <summary>
        /// Constructor. Initiated from MailStore or MailFolders object
        /// </summary>
        /// <param name="folder">Outlook.Folder object</param>
        public MailFolder(Outlook.Folder folder)
        {
            _folder = folder;
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~MailFolder()
        {
            if (_itemAddCallback != null)
            {
                UnRegisterItemAddEventHandler();
            }
        }

        /// <summary>
        /// Gets collection of subfolders within this folder
        /// </summary>
        /// <returns>MailFolders object or null if this folder has no subfolders</returns>
        public MailFolders GetSubFolders()
        {
            if (null == _folder.Folders)
            {
                return null;
            }

            return new MailFolders(_folder.Folders);
        }

        /// <summary>
        /// Adds folder as a subfolder of current folder
        /// </summary>
        /// <param name="name">Nam of new folder</param>
        /// <returns></returns>
        public MailFolder AddSubFolder(string name)
        {
            return new MailFolder(_folder.Folders.Add(name) as Outlook.Folder);
        }

        /// <summary>
        /// Deletes current folder
        /// </summary>
        public void Delete()
        {
            _folder.Delete();
        }

        /// <summary>
        /// Gets collection of MailItems in current folder
        /// </summary>
        /// <returns>MailItems object or null if this folder has no subfolders</returns>
        public MailItems GetMailItems()
        {
            if (null == _folder.Items)
            {
                return null;
            }

            return new MailItems(_folder.Items);
        }
        
        /// <summary>
        /// Folder Path of this folder
        /// </summary>
        public string FolderPath
        {
            get
            {
                return _folder.FolderPath;
            }
        }

        /// <summary>
        /// Display Name of current folder
        /// </summary>
        public string Name
        {
            get
            {
                return _folder.Name;
            }
        }

        /// <summary>
        /// Number of mail items in the current folder
        /// </summary>
        public int MailItemsCount
        {
            get
            {
                return ((null == _folder.Items) ? 0 : _folder.Items.Count);
            }
        }

        /// <summary>
        /// Number of subfolders in the current folder
        /// </summary>
        public int SubFoldersCount
        {
            get
            {
                return ((null == _folder.Folders) ? 0 : _folder.Folders.Count);
            }
        }

        /// <summary>
        /// Default Message Class of current folder
        /// </summary>
        public string DefaultMessageClass
        {
            get
            {
                return _folder.DefaultMessageClass;
            }
        }

        /// <summary>
        /// Outlook folder item
        /// </summary>
        public Outlook.Folder OutlookFolderItem
        {
            get 
            {
                return _folder;
            }
        }

        /// <summary>
        /// Registers event handler for ItemAdd event for new mail in the folder (i.e. new MailSim.OL.MailItem).
        /// </summary>
        /// <param name="callback">public static void FolderEvent(MailItem mail)</param>
        public void RegisterItemAddEventHandler(FolderItemAddEvent callback)
        {
            _itemAddCallback = callback;
            _folder.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAddEvent);
        }

        /// <summary>
        /// Unregisters event handler previously registered with RegisterItemAddEventHandler
        /// </summary>
        public void UnRegisterItemAddEventHandler()
        {
            if (_itemAddCallback != null)
            {
                _folder.Items.ItemAdd -= new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAddEvent);
                _itemAddCallback = null;
            }
        }

        /// <summary>
        /// Adds event to the item
        /// </summary>
        /// <param name="Item"></param>
        private void Items_ItemAddEvent(object Item)
        {
            if ((_itemAddCallback != null) && (Item != null) && (Item is Outlook.MailItem))
            {
                Outlook.MailItem mail = (Outlook.MailItem)Item;
                _itemAddCallback(new MailItem(mail));
            }
        }
    }
}
