using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MailSim.Common.Contracts;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Extensions;

namespace MailSim.ProvidersREST
{
    class AddressBookProvider : IAddressBook
    {
        private readonly ActiveDirectoryClient _adClient;

        public AddressBookProvider(ActiveDirectoryClient adClient)
        {
            _adClient = adClient;
        }

        public IEnumerable<string> GetUsers(string match, int count)
        {
            return EnumerateUsers(match, count);
        }

        public IEnumerable<string> GetDLMembers(string dLName, int count)
        {
            IReadOnlyList<IGroup> groups;

            if (string.IsNullOrEmpty(dLName))
            {
                groups = _adClient.Groups
                    .ExecuteAsync()
                    .Result
                    .CurrentPage;   // assume we are going to use the first matching group
            }
            else
            {
                groups = _adClient.Groups
                    .Where(g => g.Mail.StartsWith(dLName))
                    .ExecuteAsync()
                    .Result
                    .CurrentPage;   // assume we are going to use the first matching group
            }

            if (groups.Any() == false)
            {
                return Enumerable.Empty<string>();
            }

            var group = groups.First() as Group;
            IGroupFetcher groupFetcher = group;

            var pages = groupFetcher.Members.ExecuteAsync().Result;

            var members = GetFilteredItems(pages, count, (member) => member is User);

            return members.Select(m => (m as User).UserPrincipalName);
        }

        private IEnumerable<string> EnumerateUsers(string match, int count)
        {
            IPagedCollection<IUser> pagedUsers;

            if (string.IsNullOrEmpty(match))
            {
                pagedUsers = _adClient.Users
                    .ExecuteAsync()
                    .Result;
            }
            else
            {
                pagedUsers = _adClient.Users
                    .Where(x =>
                        x.UserPrincipalName.StartsWith(match) ||
                        x.DisplayName.StartsWith(match) ||
                        x.GivenName.StartsWith(match) ||
                        x.Surname.StartsWith(match)
                    )
                    .ExecuteAsync()
                    .Result;
            }

            var users = GetFilteredItems(pagedUsers, count, (u) => true);

            return users.Select(u => u.UserPrincipalName);
        }

        private IEnumerable<T> GetFilteredItems<T>(IPagedCollection<T> pages, int count, Func<T, bool> filter)
        {
            foreach (var item in pages.CurrentPage)
            {
                if (filter(item) && count-- > 0)
                {
                    yield return item;
                }
            }

            while (count > 0 && pages.MorePagesAvailable)
            {
                pages = pages.GetNextPageAsync().Result;

                foreach (var item in pages.CurrentPage)
                {
                    if (filter(item) && count-- > 0)
                    {
                        yield return item;
                    }
                }
            }
        }
    }
}
