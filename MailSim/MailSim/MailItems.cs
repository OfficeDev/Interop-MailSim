using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailSim.OL
{
    /// <summary>
    /// MailItems class
    /// </summary>
    public class MailItems
    {
        private Outlook.Items _mailItems;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="outlookItems">Outlook items</param>
        public MailItems(Outlook.Items outlookItems)
        {
            _mailItems = outlookItems;
        }

        /// <summary>
        /// Enumerator. Allows "foreach (MailFolder folder in MailFolders folders) {}"
        /// </summary>
        /// <returns>MailFolder object</returns>
        public IEnumerator GetEnumerator()
        {
            foreach (Outlook.MailItem item in _mailItems)
            {
                yield return new MailItem(item);
            }
        }

        /// <summary>
        /// Gets a particular mail item
        /// </summary>
        /// <param name="count">position of the mail item</param>
        /// <returns></returns>
        public MailItem Get(int count)
        {
            return (new MailItem(_mailItems[count]));
        }

        /// <summary>
        /// Gets the first mail item
        /// </summary>
        /// <returns>first mail item</returns>
        public MailItem GetFirst()
        {
            return (new MailItem(_mailItems.GetFirst()));
        }

        /// <summary>
        /// Gets the last mail item
        /// </summary>
        /// <returns>last mail item</returns>
        public MailItem GetLast()
        {
            return (new MailItem(_mailItems.GetLast()));
        }

        /// <summary>
        /// Gets the next mail item
        /// </summary>
        /// <returns>next mail item</returns>
        public MailItem GetNext()
        {
            return (new MailItem(_mailItems.GetNext()));
        }

        /// <summary>
        /// Gets the previous mail item
        /// </summary>
        /// <returns>previous mail item</returns>
        public MailItem GetPrevious()
        {
            return (new MailItem(_mailItems.GetPrevious()));
        }

        /// <summary>
        /// Count of MailItem objects in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _mailItems.Count;
            }
        }
    }
}
