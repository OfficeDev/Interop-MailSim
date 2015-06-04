using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailSim.OL
{
    public class OLAddressEntry
    {
        private Outlook.AddressEntry _addressEntry;

        public OLAddressEntry(Outlook.AddressEntry addressEntry)
        {
            _addressEntry = addressEntry;
        }


    }
}
