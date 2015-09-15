using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office365.OutlookServices;
using MailSim.Common.Contracts;
using Microsoft.OData.ProxyExtensions;
using MailSim.Common;

namespace MailSim.ProvidersREST
{
    class MailFolderProviderSDK : IMailFolder
    {
        private readonly Lazy<IFolderFetcher> _folderFetcher;
        private readonly OutlookServicesClient _outlookClient;
        private readonly string _id;
        private readonly bool _isRoot;

        internal MailFolderProviderSDK(OutlookServicesClient outlookClient, IFolder folder)
        {
            _outlookClient = outlookClient;
            Name = folder.DisplayName;

            _id = folder.Id;

            _folderFetcher = new Lazy<IFolderFetcher>(() => _outlookClient.Me.Folders.GetById(_id));
        }

        internal MailFolderProviderSDK(OutlookServicesClient outlookClient, string rootName)
        {
            _isRoot = true;

            _outlookClient = outlookClient;
            Name = rootName;
        }

        public string Name { get; private set; }

        public string FolderPath
        {
            get
            {
                return Name;    // TODO: is it the right thing to do?
            }
        }

        public int MailItemsCount
        {
            get
            {
                return GetMailCount();
            }
        }

        public int SubFoldersCount
        {
            get
            {
                if (_isRoot)
                {
                    // TODO: CountAsync() method fails; have to use direct HTTP call
#if true
                    return FoldersCountRequest().Result;
#else
                    return (int) _outlookClient.Me.Folders.CountAsync().Result;
#endif
                }
                else
                {
                    IFolder folder = _folderFetcher.Value.ExecuteAsync().Result;
                    return folder.ChildFolderCount ?? 0;
                }
            }
        }

        public IEnumerable<IMailItem> MailItems
        {
            get
            {
                // Generate a request that gets all mail items
                return GetMailItems(string.Empty, GetMailCount());
            }
        }

        public IEnumerable<IMailItem> GetMailItems(string filter, int count)
        {
            // TODO: there is no way right now to filter mails server-side
            var pages = _folderFetcher.Value.Messages
                .Take(100)      // set the page size
                .ExecuteAsync()
                .Result;

            filter = filter ?? string.Empty;

            var items = GetFilteredItems(pages, count, (i) => i.Subject.ContainsCaseInsensitive(filter));

            return items.Select(i => new MailItemProviderSDK(_outlookClient, i));
        }

        public void Delete()
        {
            var folder = _folderFetcher.Value.ExecuteAsync().Result;

            folder.DeleteAsync().Wait();
        }

        public IMailFolder AddSubFolder(string name)
        {
            Folder newFolder = new Folder
            {
                DisplayName = name
            };

            _folderFetcher.Value.ChildFolders.AddFolderAsync(newFolder).Wait();

            return new MailFolderProviderSDK(_outlookClient, newFolder);
        }

        public IEnumerable<IMailFolder> SubFolders
        {
            get
            {
                return GetSubFolders();
            }
        }

        internal string Handle
        {
            get
            {
                return _id;
            }
        }
 
        public void RegisterItemAddEventHandler(Action<IMailItem> callback)
        {
            // TODO: Implement this
        }

        public void UnRegisterItemAddEventHandler()
        {
            // TODO: Implement this
        }

        private IEnumerable<T> GetFilteredItems<T>(IPagedCollection<T> pages, int count, Func<T, bool> filter)
        {
            foreach (var item in pages.CurrentPage)
            {
                if (--count < 0)
                {
                    yield break;
                }
                else if (filter(item))
                {
                    yield return item;
                }
            }

            while (count > 0 && pages.MorePagesAvailable)
            {
                pages = pages.GetNextPageAsync().Result;

                foreach (var item in pages.CurrentPage)
                {
                    if (--count < 0)
                    {
                        yield break;
                    }
                    else if (filter(item))
                    {
                        yield return item;
                    }
                }
            }
        }

        private int GetMailCount()
        {
            long count = 0;
            // TODO: CountAsync() method fails; have to use direct HTTP call
#if true
            count = MailCountRequest(_id).Result;
#else
            count = _folderFetcher.Messages.CountAsync().Result;
#endif
            return (int)count;
        }

        private IEnumerable<IMailFolder> GetSubFolders()
        {
            var folderCollection = _isRoot ? _outlookClient.Me.Folders : _folderFetcher.Value.ChildFolders;

            IPagedCollection<IFolder> folders = folderCollection.ExecuteAsync().Result;

            var allFolders = GetFilteredItems(folders, int.MaxValue, (f) => true);

            return allFolders.Select(f => new MailFolderProviderSDK(_outlookClient, f));
        }

        private async Task<int> MailCountRequest(string folderId)
        {
            string uri = string.Format("Folders/{0}/Messages/$count", folderId);

            return await HttpUtil.GetItemAsync<int>(uri);
        }

        private async Task<int> FoldersCountRequest()
        {
            return await HttpUtil.GetItemAsync<int>("Folders/$count");
        }
    }
}
