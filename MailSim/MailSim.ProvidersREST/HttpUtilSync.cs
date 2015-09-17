using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSim.ProvidersREST
{
    static class HttpUtilSync
    {
        internal static T GetItem<T>(string uri)
        {
            return HttpUtil.GetItemAsync<T>(uri).GetResult();
        }

        internal static T GetItems<T>(string uri)
        {
            return HttpUtil.GetItemsAsync<T>(uri).GetResult();
        }

        internal static T PostItem<T>(string uri, T item = default(T))
        {
            return HttpUtil.PostItemAsync<T>(uri, item).GetResult();
        }

        internal static T PostItemDynamic<T>(string uri, dynamic body)
        {
            // Can't use extensions with dynamic types...
            return HttpUtil.PostItemDynamicAsync<T>(uri, body).ConfigureAwait(false)
                .GetAwaiter()
                .GetResult();
        }

        internal static void DeleteItem(string uri)
        {
            HttpUtil.DeleteItemAsync(uri).GetResult();
        }

        internal static T PatchItem<T>(string uri, T item)
        {
            return HttpUtil.PatchItemAsync<T>(uri, item).GetResult();
        }
    }
}

