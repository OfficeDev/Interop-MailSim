using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSim.Contracts
{
    interface IMailItem
    {
        /// <summary>
        /// Mail subject field
        /// </summary>
        string Subject { get; set; }
        /// <summary>
        /// Mail body
        /// </summary>
        string Body { get; set; }
        /// <summary>
        /// Mail sender name
        /// </summary>
        string SenderName { get; }
        /// <summary>
        /// Adds recipient to recipient's list of the message
        /// </summary>
        /// <param name="recipient">recipient</param>
        void AddRecipient(string recipient);
        /// <summary>
        /// Adds file attachment to the current message
        /// </summary>
        /// <param name="file">Full file path to the file to attach</param>
        void AddAttachment(string fileName);
        /// <summary>
        /// Adds message as attachment to current message
        /// </summary>
        /// <param name="mailitem"></param>
        void AddAttachment(IMailItem mailItem);
        /// <summary>
        /// Send this message
        /// </summary>
        void Send();
        /// <summary>
        /// Deletes this message
        /// </summary>
        void Delete();
        /// <summary>
        /// Moves Mail item into new folder
        /// </summary>
        /// <param name="destination" - IMailFolder representing folder to move to></param>
        void Move(IMailFolder destination);
        /// <summary>
        /// Resolves and validates all recipients. Returns true if successful; false if one or more recipients cannot be resolved.  
        /// </summary>
        bool ValidateRecipients();
        /// <summary>
        /// Creates a reply, pre-addressed to the original sender or all original recipients, from the original message
        /// </summary>
        /// <param name="replyAll" - replies to all original recipients if true; only to the original sender if false></param>
        /// <returns>IMailItem object that represents the reply</returns>
        IMailItem Reply(bool replyAll);
        /// <summary>
        /// Executes the Forward action for an item and returns the resulting copy as a MailItem object
        /// </summary>
        /// <returns>IMailItem object that represents the new mail item</returns>
        IMailItem Forward();
#if false
        /// <summary>
        /// HTML Body of the email
        /// </summary>
        public string HTMLBody { get; set; }
#endif
    }
}
