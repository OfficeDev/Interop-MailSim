//--------------------------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized
// to use this sample source code. For the terms of the license, please see the
// license agreement between you and Microsoft.
//--------------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailSim.OL
{
    public class MailItem
    {
        private Outlook._MailItem _mailitem;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="mailItem"></param>
        public MailItem(Outlook.MailItem mailItem)
        {
            _mailitem = mailItem;
        }

        /// <summary>
        /// Adds recipient to recipient's list of the message
        /// </summary>
        /// <param name="recipient">recipient</param>
        public void AddRecipient(string recipient)
        {
            _mailitem.Recipients.Add(recipient);
            return;
        }

        /// <summary>
        /// Adds file attachment to the current message
        /// </summary>
        /// <param name="file">Full file path to the file to attach</param>
        public void AddAttachment(string file)
        {
            _mailitem.Attachments.Add(file);
            return;
        }

        /// <summary>
        /// Adds message as attachment to current message
        /// </summary>
        /// <param name="mailitem"></param>
        public void AddAttachment(MailItem mailitem)
        {
            _mailitem.Attachments.Add(mailitem.OutlookMailItem);
            return;
        }

        /// <summary>
        /// Send current message (MailItem)
        /// </summary>
        public void Send()
        {
            _mailitem.Send();
        }

        /// <summary>
        /// Creates a reply, pre-addressed to the original sender or all original recipients, from the original message
        /// </summary>
        /// <param name="replyAll" - replies to all original recipients if true; only to the original sender if false></param>
        /// <returns>A MailItem object that represents the reply</returns>
        public MailItem Reply(bool replyAll)
        {
            if (replyAll)
            {
                return new MailItem(_mailitem.ReplyAll());
            }
            else
            {
                return new MailItem(_mailitem.Reply());
            }
        }

        /// <summary>
        /// Executes the Forward action for an item and returns the resulting copy as a MailItem object
        /// </summary>
        /// <returns>A MailItem object that represents the new mail item</returns>
        public MailItem Forward()
        {
            return new MailItem(_mailitem.Forward());
        }

        /// <summary>
        /// Deletes current mailitem
        /// </summary>
        public void Delete()
        {
            _mailitem.Delete();
        }

        /// <summary>
        /// Moves Mail item into new folder
        /// </summary>
        /// <param name="newFolder" - Folder object representing Folder to move to></param>
        public void Move(MailFolder newFolder)
        {
            _mailitem = _mailitem.Move(newFolder.OutlookFolderItem);
        }

        /// <summary>
        /// Mail subject field
        /// </summary>
        public string Subject
        {
            get
            {
                return _mailitem.Subject;
            }

            set
            {
                _mailitem.Subject = value;
            }
        }

        /// <summary>
        /// Mail body
        /// </summary>
        public string Body
        {
            get
            {
                return _mailitem.Body;
            }

            set
            {
                _mailitem.Body = value;
            }
        }
        /// <summary>
        /// HTML Body of the email
        /// </summary>
        public string HTMLBody
        {
            get
            {
                return _mailitem.HTMLBody;
            }

            set
            {
                _mailitem.HTMLBody = value;
            }
        }
        
        /// <summary>
        /// Mail sender name
        /// </summary>
        public string SenderName
        {
            get
            {
                return _mailitem.Sender.Name;
            }
        }

        /// <summary>
        /// Resolves and validates all recipients. Returns true if successful; false if one or more recipients cannot be resolved.  
        /// </summary>
        public bool ValidateRecipients
        {
            get
            {
                return (_mailitem.Recipients.ResolveAll());
            }
        }

        public Outlook.MailItem OutlookMailItem
        {
            get
            {
                return (Outlook.MailItem) _mailitem;
            }
        }
    }
}
