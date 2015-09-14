using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Outlook;
using MailSim.Common;
using MailSim.Common.Contracts;

namespace MailSim.ProvidersOM
{
    class AddressBookProviderOM : IAddressBook
    {
       private readonly AddressList _addressList;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="addressList"></param>
        public AddressBookProviderOM(AddressList addressList)
        {
            _addressList = addressList;
        }

        public IEnumerable<string> GetUsers(string match, int count)
        {
            match = match ?? string.Empty;

            foreach (AddressEntry addrEntry in _addressList.AddressEntries)
            {
                if (addrEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry)
                {
                    if (--count < 0)
                    {
                        yield break;
                    }
                    else if (addrEntry.Name.ContainsCaseInsensitive(match))
                    {
                        yield return addrEntry.GetExchangeUser().PrimarySmtpAddress;
                    }
                }
            }
        }

        public IEnumerable<string> GetDLMembers(string dLName, int count)
        {
            if (string.IsNullOrEmpty(dLName))
            {
                yield break;
            }

            foreach (AddressEntry addrEntry in _addressList.AddressEntries)
            {
                if (addrEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeDistributionListAddressEntry
                    && addrEntry.Name.EqualsCaseInsensitive(dLName))
                {
                    foreach (AddressEntry member in addrEntry.GetExchangeDistributionList().GetExchangeDistributionListMembers())
                    {
                        if (--count < 0)
                        {
                            yield break;
                        }
                        if (member.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry)
                        {
                            yield return member.GetExchangeUser().PrimarySmtpAddress;
                        }
                        else if (addrEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                        {
                            yield return member.GetExchangeDistributionList().PrimarySmtpAddress;
                        }
                    }
               }
            }
        }
    }
}
