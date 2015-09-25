using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common.Contracts;
using System.Dynamic;
using MailSim.Common;

namespace MailSim.ProvidersREST
{
    class MailFolderProviderHTTP : HTTP.BaseProviderHttp, IMailFolder
    {
        private const int PageSize = 100;   // the page to use for $top argument
        private readonly Folder _folder;
//        private string _subscriptionId;

        internal MailFolderProviderHTTP(Folder folder, string name = null)
        {
            _folder = folder;

            if (name != null)
            {
                Name = name;
            }
            else if (_folder != null)
            {
                Name = _folder.DisplayName;
            }
        }

        public string Name { get; private set; }

        public string FolderPath
        {
            get
            {
                return Name;
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
                return GetChildFolderCount();
            }
        }

        public IEnumerable<IMailFolder> SubFolders
        {
            get
            {
                return GetSubFolders();
            }
        }

        public IEnumerable<IMailItem> MailItems
        {
            get
            {
                return GetMailItems(string.Empty, GetMailCount());
            }
        }

        public void Delete()
        {
            HttpUtilSync.DeleteItem(Uri);
        }

        public IEnumerable<IMailItem> GetMailItems(string filter, int count)
        {
            var msgs = GetMessages(count);

            var items = msgs.Select(x => new MailItemProviderHTTP(x));

            filter = filter ?? string.Empty;

            return items.Where(i => i.Subject.ContainsCaseInsensitive(filter));
        }

        public IMailFolder AddSubFolder(string name)
        {
            dynamic folderName = new ExpandoObject();
            folderName.DisplayName = name;

            Folder newFolder = HttpUtilSync.PostItemDynamic<Folder>(Uri + "/ChildFolders", folderName);

            return new MailFolderProviderHTTP(newFolder);
        }

        // TODO: Implement this after Notifications graduate from preview state
        public void RegisterItemAddEventHandler(Action<IMailItem> callback)
        {
#if false
            string baseUri = "https://outlook.office.com/api/beta/me";

            string uri = baseUri + "/subscriptions";

            var res = Util.DoHttp<SubscriptionRequest, SubscriptionResponse>("POST", uri, new SubscriptionRequest()
            {
                ResourceURL = string.Format("{0}/{1}/messages", baseUri, Uri),
                Type = "#Microsoft.OutlookServices.PushSubscription",
                CallbackURL = "https://webhook.azurewebsites.net/api/send/myNotifyClient",
                ChangeType = "Created",
                ClientState = "3250be24-1282-4b46-a41e-0e53b4cae73f"    // GUID
            }).Result;

            _subscriptionId = res.Id;

            StartNotificationListener(res.Id, callback);
#endif
        }

        // TODO: Implement this after Notifications graduate from preview state
        public void UnRegisterItemAddEventHandler()
        {
#if false
            string baseUri = "https://outlook.office.com/api/beta/me";

            string uri = string.Format("{0}/subscriptions('{1}')", baseUri, _subscriptionId);

            Util.DeleteAsync(uri).Wait();
#endif
        }

        internal string Handle
        {
            get
            {
                return _folder.Id;
            }
        }

        private string Uri
        {
            get
            {
                return string.Format("Folders/{0}", _folder.Id);
            }
        }

        private int GetMailCount()
        {
            return HttpUtilSync.GetItem<int>(Uri + "/Messages/$count");
        }

        private int GetChildFolderCount()
        {
            if (_folder == null)
            {
                return HttpUtilSync.GetItem<int>("Folders/$count");
            }
            else
            {
                return HttpUtilSync.GetItem<int>(Uri + "/ChildFolders/$count");
            }
        }

        private IEnumerable<MailItemProviderHTTP.Message> GetMessages(int count)
        {
            string uri = Uri + string.Format("/Messages?&$top={0}", PageSize);

            return HttpUtilSync.GetItems<MailItemProviderHTTP.Message>(uri, count);
        }

        private static void StartNotificationListener(string id, Action<IMailItem> callback)
        {
        }

        private IEnumerable<IMailFolder> GetSubFolders()
        {
            string uri = _folder == null ? "Folders" : Uri + "/ChildFolders";

            var folders = HttpUtilSync.GetItems<Folder>(uri, int.MaxValue);

            return folders.Select(f => new MailFolderProviderHTTP(f));
        }

        internal class Folder
        {
            public string Id { get; set; }
            public string DisplayName { get; set; }
            public int ChildFolderCount { get; set; }
        }

        private class SubscriptionResponse
        {
            public string Id { get; set; }
            public string ChangeType { get; set; }
            public DateTime ExpirationTime { get; set; }
        }

        private class SubscriptionRequest
        {
            [Newtonsoft.Json.JsonProperty("@odata.type")]
            public string Type { get; set; }
            public string ResourceURL { get; set; }
            public string CallbackURL { get; set; }
            public string ChangeType { get; set; }
            public string ClientState { get; set; }
        }
    }
}
