using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common.Contracts;
using System.IO;

namespace MailSim.ProvidersREST
{
    class MailItemProviderSDK : IMailItem
    {
        private readonly OutlookServicesClient _outlookClient;
        private readonly IMessage _message;

        public MailItemProviderSDK(OutlookServicesClient outlookClient, IMessage msg)
        {
            _outlookClient = outlookClient;
           _message = msg;
        }

        public string Subject
        {
            get
            {
                return _message.Subject;
            }

            set
            {
                SetAndUpdate((message) => message.Subject = value);
            }
        }

        public string Body
        {
            get
            {
                return _message.Body.Content;
            }

            set
            {
                SetAndUpdate((message) =>
                    message.Body = new ItemBody
                    {
                        Content = value,
                        ContentType = BodyType.HTML
                    });
            }
        }
 
        public string SenderName
        {
            get
            {
                return _message.Sender.EmailAddress.Address;
            }
        }

        public void AddRecipient(string recipient)
        {
            SetAndUpdate((message) => message.ToRecipients.Add(new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient
                }
            }));
        }

        public void AddAttachment(string filepath)
        {
            using (var reader = new StreamReader(filepath))
            {
                var contents = reader.ReadToEnd();

                var msgFetcher = _outlookClient.Me.Messages.GetById(_message.Id);

                var bytes = System.Text.Encoding.Unicode.GetBytes(contents);
                var name = filepath.Split('\\').Last();

                var fileAttachment = new FileAttachment
                {
                    ContentBytes = bytes,
                    Name = name,
                    Size = bytes.Length
                };

                msgFetcher.Attachments.AddAttachmentAsync(fileAttachment).GetResult();
            }
        }

        // TODO: Figure out how to implement this
        public void AddAttachment(IMailItem mailItem)
        {
#if false
            var itemProvider = mailItem as MailItemProvider;

            var msgFetcher = _outlookClient.Me.Messages.GetById(_message.Id);

            var itemAttachment = new ItemAttachment
            {
                Item = itemProvider.Handle as Message,

                Name = "Item Attachment!!!",
            };

           Util.Synchronize(async () => await msgFetcher.Attachments.AddAttachmentAsync(itemAttachment));
#endif
        }

        // Create a reply message
        public IMailItem Reply(bool replyAll)
        {
            IMessage replyMsg = null;

            if (replyAll)
            {
                replyMsg = Message.CreateReplyAllAsync().GetResult();
            }
            else
            {
                replyMsg = Message.CreateReplyAsync().GetResult();
            }

            return new MailItemProviderSDK(_outlookClient, replyMsg);
        }

        public IMailItem Forward()
        {
            IMessage msg = null;

            msg = Message.CreateForwardAsync().GetResult();

            return new MailItemProviderSDK(_outlookClient, msg);
        }

        public void Send()
        {
            // This generates Me/SendMail; all the data should be in the body
            Message.SendAsync().GetResult();
        }
        
        // TODO: Should this method return a IMailItem?
        public void Move(IMailFolder newFolder)
        {
            var folderProvider = newFolder as MailFolderProviderSDK;

            var folderId = folderProvider.Handle;
            Message.MoveAsync(folderId).GetResult();
        }

        public void Delete()
        {
            _message.DeleteAsync().GetResult();
        }

        public bool ValidateRecipients()
        {
            // TODO: Implement this
            return true;
        }

        internal Item Handle
        {
            get
            {
                return _message as Item;
            }
        }

        private IMessageFetcher Message
        {
            get
            {
                return _outlookClient.Me.Messages[_message.Id];
            }
        }

        private void SetAndUpdate(Action<IMessage> action)
        {
            IMessage message = null;

            message = _outlookClient.Me.Messages[_message.Id].ExecuteAsync().GetResult();
            action(message);

            message.UpdateAsync().GetResult();
        }
    }
}
