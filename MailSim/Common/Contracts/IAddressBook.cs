using System.Collections.Generic;

namespace MailSim.Common.Contracts
{
    public interface IAddressBook
    {
        /// <summary>
        /// Builds list of addresses for all users in the Address List that have display name match
        /// </summary>
        /// <param name="match"> string to match in user name or null to return all users in the GAL</param>
        /// <param name="count"> maximum number of users to fetch from GAL</param> 
        /// <returns>List of SMTP addresses of matching users in the address list. The list will be empty if no users exist or match.</returns>
        IEnumerable<string> GetUsers(string match, int count);

        /// <summary>
        /// Builds list of addresses for all members of Exchange Distribution list in the Address List
        /// </summary>
        /// <param name="dLName">Exchange Distribution List Name</param>
        /// <param name="count"> maximum number of DL members to fetch from GAL</param> 
        /// <returns>List of SMTP addresses of DL members or null if DL is not found. Nesting DLs are not expanded. </returns>
        IEnumerable<string> GetDLMembers(string dLName, int count);
    }
}
