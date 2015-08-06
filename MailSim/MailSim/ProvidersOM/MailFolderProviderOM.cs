using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using MailSim.Contracts;

namespace MailSim.ProvidersOM
{
    class MailFolderProviderOM : IMailFolder 
    {
        private readonly Outlook.Folder _folder;
        private Action<IMailItem> _itemAddCallback;

        internal MailFolderProviderOM(Outlook.Folder folder)
        {
            _folder = folder;
        }

        public string Name
        {
            get
            {
                return _folder.Name;
            }
        }

        public string FolderPath
        {
            get
            {
                return _folder.FolderPath;
            }
        }

        public IEnumerable<IMailFolder> SubFolders
        {
            get
            {
                return GetSubFolders();
            }
        }

        private IEnumerable<IMailFolder> GetSubFolders()
        {
            foreach (var f in _folder.Folders)
            {
                yield return new MailFolderProviderOM(f as Outlook.Folder);
            }
        }

        public void Delete()
        {
            _folder.Delete();
        }

        public void RegisterItemAddEventHandler(Action<IMailItem> callback)
        {
            _itemAddCallback = callback;
            _folder.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAddEvent);
        }

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
                _itemAddCallback(new MailItemProviderOM(mail));
            }
        }

        public int MailItemsCount
        {
            get
            {
                return ((null == _folder.Items) ? 0 : _folder.Items.Count);
            }
        }
        public int SubFoldersCount
        {
            get
            {
                return ((null == _folder.Folders) ? 0 : _folder.Folders.Count);
            }
        }

        public IMailFolder AddSubFolder(string name)
        {
            return new MailFolderProviderOM(_folder.Folders.Add(name) as Outlook.Folder);
        }

        public IEnumerable<IMailItem> MailItems
        {
            get
            {
                return GetMailItems();
            }
        }

        private IEnumerable<IMailItem> GetMailItems()
        {
            if (null == _folder.Items)
            {
                yield break;
            }

            foreach (var item in _folder.Items)
            {
                yield return new MailItemProviderOM(item as Outlook.MailItem);
            }
        }

        internal Outlook.Folder Handle
        {
            get
            {
                return _folder;
            }
        }
    }
}
