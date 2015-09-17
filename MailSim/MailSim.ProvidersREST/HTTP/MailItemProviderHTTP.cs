using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MailSim.Common.Contracts;
using System.IO;
using System.Dynamic;

namespace MailSim.ProvidersREST
{
    class MailItemProviderHTTP : IMailItem
    {
        private Message _message;

        public MailItemProviderHTTP(Message msg)
        {
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
                _message.Subject = value;
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
                _message.Body = new ItemBody
                {
                    Content = value,
                    ContentType = "HTML"
                };
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
            _message.ToRecipients.Add(new Recipient
            {
                EmailAddress = new EmailAddress
                {
                    Address = recipient
                }
            });
        }

        public void AddAttachment(string filepath)
        {
            using (var reader = new StreamReader(filepath))
            {
                var contents = reader.ReadToEnd();

                var bytes = System.Text.Encoding.Unicode.GetBytes(contents);
                var name = filepath.Split('\\').Last();

                var fileAttachment = new FileAttachment
                {
                    ContentBytes = bytes,
                    Name = name,
                };

                HttpUtilSync.PostItem(Uri + "/attachments", fileAttachment);
            }
        }

        // TODO: Figure out how to implement this
        public void AddAttachment(IMailItem mailItem)
        {
#if false
            var itemProvider = mailItem as MailItemProvider;

            var itemAttachment = new ItemAttachment
            {
                Item = new Item { Id = itemProvider.Handle.Id },
                Name = "Item Attachment!!!",
            };

           Util.PostAsync<ItemAttachment>(Uri + "/attachments", itemAttachment).Wait();
#endif
        }

        // Create a reply message
        public IMailItem Reply(bool replyAll)
        {
            string uri = Uri + (replyAll ? "/CreateReplyAll" : "/CreateReply");

            Message replyMsg = HttpUtilSync.PostItem<Message>(Uri + "/CreateReply");

            return new MailItemProviderHTTP(replyMsg);
        }

        public IMailItem Forward()
        {
            Message msg = HttpUtilSync.PostItem<Message>(Uri + "/CreateForward");

            return new MailItemProviderHTTP(msg);
        }

        public void Send()
        {
            HttpUtilSync.PatchItem(Uri, _message);
            HttpUtilSync.PostItem<Message>(Uri + "/Send");
        }

        // TODO: Should this method return a IMailItem?
        public void Move(IMailFolder newFolder)
        {
            var folderProvider = newFolder as MailFolderProviderHTTP;

            var folderId = folderProvider.Handle;

            dynamic destination = new ExpandoObject();
            destination.DestinationId = folderId;

            HttpUtilSync.PostItemDynamic<Message>(Uri + "/Move", destination);
        }

        public void Delete()
        {
            HttpUtilSync.DeleteItem(Uri);
            _message = null;
        }

        public bool ValidateRecipients()
        {
            // TODO: Implement this
            return true;
        }

        #if false
        internal Message Handle
        {
            get
            {
                return _message;
            }
        }
#endif
        private string Uri
        {
            get
            {
                return string.Format("/Messages/{0}", _message.Id);
            }
        }

        internal class Message
        {
            public string Id { get; set; }
            public string Subject { get; set; }
            public ItemBody Body { get; set; }
            public ICollection<Recipient> ToRecipients { get; set; }
            public Sender Sender { get; set; }
            public string Importance { get; set; }
        }

        internal class ItemBody
        {
            public string Content { get; set; }
            public string ContentType { get; set; }
        }

        internal class Recipient
        {
            public EmailAddress EmailAddress { get; set; }
        }

        internal class EmailAddress
        {
            public string Address { get; set; }
            public string Name { get; set; }
        }

        internal class Sender
        {
            public EmailAddress EmailAddress { get; set; }
        }

        internal class FileAttachment
        {
            [Newtonsoft.Json.JsonProperty("@odata.type")]
            public string Type
            {
                get
                {
                    return "#Microsoft.OutlookServices.FileAttachment";
                }
            }

            public byte[] ContentBytes { get; set; }
            public string Name { get; set; }
        }

        internal class ItemAttachment
        {
            [Newtonsoft.Json.JsonProperty("@odata.type")]
            public string Type
            {
                get
                {
                    return "#Microsoft.OutlookServices.ItemAttachment";
                }
            }
            public string Name { get; set; }
            public Item Item { get; set; }
        }

        internal class Item
        {
            public string Id { get; set; }
        }
    }
}
