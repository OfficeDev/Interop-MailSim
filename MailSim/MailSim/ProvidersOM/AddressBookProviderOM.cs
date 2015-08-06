using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using MailSim.Contracts;

namespace MailSim.ProvidersOM
{
    class AddressBookProviderOM : IAddressBook
    {
       private readonly Outlook.AddressList _addressList;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="addressList"></param>
        public AddressBookProviderOM(Outlook.AddressList addressList)
        {
            _addressList = addressList;
        }

        /// <summary>
        /// Builds list of addresses for all users in the Address List that have display name match
        /// </summary>
        /// <param name="match"> string to match in user name or null to return all users in the GAL</param>
        /// <returns>List of SMTP addresses of matching users in the address list. The list will be empty if no users exist or match.</returns>
        public IEnumerable<string> GetUsers(string match)
        {
            foreach (Outlook.AddressEntry addrEntry in _addressList.AddressEntries)
            {
                if (addrEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                {
                    if ((match == null) || addrEntry.Name.Contains(match))
                    {
                        yield return addrEntry.GetExchangeUser().PrimarySmtpAddress;
                    }
                }
            }
        }

        /// <summary>
        /// Builds list of addresses for all members of Exchange Distribution list in the Address List
        /// </summary>
        /// <param name="dLName">Exchane Distribution List Name</param>
        /// <returns>List of SMTP addresses of DL members or null if DL is not found. Nesting DLs are not expanded. </returns>
        public IEnumerable<string> GetDLMembers(string dLName)
        {
            foreach (Outlook.AddressEntry addrEntry in _addressList.AddressEntries)
            {
                if ((addrEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                    && (addrEntry.Name.Equals(dLName, StringComparison.OrdinalIgnoreCase)))
                {
                    foreach(Outlook.AddressEntry member in addrEntry.GetExchangeDistributionList().GetExchangeDistributionListMembers())
                    {
                        if (member.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                        {
                            yield return member.GetExchangeUser().PrimarySmtpAddress;
                        }
                        else if (addrEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                        {
                            yield return member.GetExchangeDistributionList().PrimarySmtpAddress;
                        }
                    }
               }
            }
        }
    }
}
