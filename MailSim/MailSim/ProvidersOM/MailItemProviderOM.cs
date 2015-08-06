using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;
using MailSim.Contracts;

namespace MailSim.ProvidersOM
{
    class MailItemProviderOM : IMailItem
    {
        private Outlook._MailItem _mailItem;

        internal MailItemProviderOM(Outlook.MailItem mailItem)
        {
            _mailItem = mailItem;
        }

        public void AddRecipient(string recipient)
        {
            _mailItem.Recipients.Add(recipient);
        }

        public void AddAttachment(string file)
        {
            _mailItem.Attachments.Add(file);
        }

        public void AddAttachment(IMailItem mailItem)
        {
            var provider = mailItem as MailItemProviderOM;
            _mailItem.Attachments.Add(provider.Handle);
        }

        public void Move(IMailFolder newFolder)
        {
            var provider = newFolder as MailFolderProviderOM;
            _mailItem = _mailItem.Move(provider.Handle);
        }

        public void Delete()
        {
            _mailItem.Delete();
        }

        public void Send()
        {
            _mailItem.Send();
        }

        public IMailItem Reply(bool replyAll)
        {
            if (replyAll)
            {
                return new MailItemProviderOM(_mailItem.ReplyAll());
            }
            else
            {
                return new MailItemProviderOM(_mailItem.Reply());
            }
        }

        public IMailItem Forward()
        {
            return new MailItemProviderOM(_mailItem.Forward());
        }

        public string Subject
        {
            get
            {
                return _mailItem.Subject;
            }

            set
            {
                _mailItem.Subject = value;
            }
        }

        public string Body
        {
            get
            {
                return _mailItem.Body;
            }

            set
            {
                _mailItem.Body = value;
            }
        }

        public string SenderName
        {
            get
            {
                return _mailItem.Sender.Name;
            }
        }

        internal Outlook._MailItem Handle
        {
            get
            {
                return _mailItem;
            }
        }

        public bool ValidateRecipients()
        {
            return (_mailItem.Recipients.ResolveAll());
        }
    }
}
