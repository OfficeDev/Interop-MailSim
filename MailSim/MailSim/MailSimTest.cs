//--------------------------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized
// to use this sample source code. For the terms of the license, please see the
// license agreement between you and Microsoft.
//--------------------------------------------------------------------------------------


//--------------------------------------------------------------------------------------
//
// NOTE: 
// THIS IS A FILE USED FOR FEATURE DEVELOPMENT AND TESTING, THIS IS NOT USED FOR 
// NORMAL MAILSIM OPERATION.
//
//--------------------------------------------------------------------------------------


using System;
using System.Collections.Generic;
using System.Linq;

using MailSim.Contracts;

namespace MailSim
{
    class MailSimTest
    {
        /// <summary>
        /// Test module, focusing on Outlook (OOM) wrapper classes (Mail*)
        /// Also serves as an example for Mail* classes usage 
        /// </summary>
        /// <param name="args">Command line argument. It is expected that the first argument is always "/t"</param>
        public void Execute(string[] args)
        {
            IMailStore mailStore = null;

            try
            {
                // We will use mailbox with display name specified in arg[1];
                // otherwise, we will get the default store
                mailStore = ProviderFactory.CreateMailStore(args.Length > 1 ? args[1] : null);
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }

            string mailboxName = mailStore.DisplayName;

            // Display all top folders in the mailbox
            IMailFolder rootFolder = mailStore.RootFolder;

            var rootFolders = rootFolder.SubFolders;

            foreach (var mailFolder in rootFolders)
            {
                Console.WriteLine(mailFolder.FolderPath);
            }

            // Open Inbox and loop through it's top folders
            IMailFolder inbox = mailStore.GetDefaultFolder("olFolderInbox");

            // Registering callback for Inbox "ItemAdd" event
            inbox.RegisterItemAddEventHandler(FolderEvent);
            
            Console.WriteLine("Inbox has {0} mail items", inbox.MailItemsCount);
            var inboxSubFolders = inbox.SubFolders;
            IMailFolder testFolder = null;
            const string testFolderName = "MailSim Test Folder";
            Console.WriteLine("Exploring Inbox Folders");

            foreach (IMailFolder mailFolder in inboxSubFolders)
            {
                if (testFolderName == mailFolder.Name)
                {
                    testFolder = mailFolder;
                }
                Console.WriteLine(mailFolder.FolderPath);
            }

            if (null == testFolder)
            {
                testFolder = inbox.AddSubFolder(testFolderName);
            }

            // Adding folder under Test Folder
            string folderName = "Test Subfolder " + (testFolder.SubFoldersCount + 1).ToString() + " - " + DateTime.Now.TimeOfDay.ToString();
            testFolder.AddSubFolder(folderName);

            // Deleting folder under Test Folder
            if (testFolder.SubFoldersCount > 2)
            {
                testFolder.SubFolders.First().Delete();
            }

            // Deleting email in inbox
            if (inbox.MailItemsCount > 2)
            {
                var mailToDelete = inbox.MailItems.First();
                Console.WriteLine("Deleting email {0}", mailToDelete.Subject);
                mailToDelete.Delete();
            }

            // Moving email 
            if (inbox.MailItemsCount > 2)
            {
                inbox.MailItems.First().Move(testFolder);
            }
            
            // Sending new email with message attachment to matching users
            Console.WriteLine("Sending new email to matching GAL users");
            IMailItem newMail = mailStore.NewMailItem();
            newMail.Subject = "Test Mail from MailSim to GAL users " + (inbox.MailItemsCount + 1).ToString() + " - " + DateTime.Now.TimeOfDay.ToString();
            newMail.Body = "Test from MailSim to matching users";

            if (inbox.MailItemsCount > 0)
            {
                var message = inbox.MailItems.First();
                newMail.AddAttachment(message);
            }
            newMail.AddRecipient(mailboxName);

            var gal = mailStore.GetGlobalAddressList();
            if (gal != null)
            {
                var users = gal.GetUsers("Mailsim");
                foreach (string userAddress in users)
                {
                    newMail.AddRecipient(userAddress);
                }
            }

            if (newMail.ValidateRecipients())
            {
                newMail.Send();
                Console.WriteLine("Mail to specified users  sent!");
            }
            else
            {
                Console.WriteLine("Incorrect recipient(s), mail not sent");
            }

            // Sending new email with to DL members
            Console.WriteLine("Sending new email to DL users");
            newMail = mailStore.NewMailItem();
            newMail.Subject = "Test Mail from MailSim to DL members " + (inbox.MailItemsCount + 1).ToString() + " - " + DateTime.Now.TimeOfDay.ToString();
            newMail.Body = "Test from MailSim to DL members";
            newMail.AddRecipient(mailboxName);

            if (gal != null)
            {
                var members = gal.GetDLMembers("Mailsim Users");

                if (members.Any() == false)
                {
                    Console.WriteLine("ERROR: Distribution list not found or empty");
                }
                else
                {
                    foreach (string userAddress in members)
                    {
                        newMail.AddRecipient(userAddress);
                    }
                }
            }

            if (newMail.ValidateRecipients())
            {
                newMail.Send();
                Console.WriteLine("Mail to distribution list members sent!");
            }
            else
            {
                Console.WriteLine("Incorrect recipient(s), mail not sent");
            }


            var inboxMailItems = inbox.MailItems;
            // Reply All
            if (inbox.MailItemsCount >= 1)
            {
                var replyMail = inboxMailItems.First().Reply(true);
                replyMail.Body = "Reply All by MailSim" + replyMail.Body;
                Console.WriteLine("Message Body:");
                Console.WriteLine(replyMail.Body);
                replyMail.Send();
                Console.WriteLine("Reply All mail sent!");
            }

            // Forward
            if (inbox.MailItemsCount >= 2)
            {
                var forwardMail = inboxMailItems.Skip(1).First().Forward();
                forwardMail.Body = "Forward by MailSim" + forwardMail.Body;
                Console.WriteLine("Message Body:");
                Console.WriteLine(forwardMail.Body);
                forwardMail.AddRecipient(mailboxName);
                if (forwardMail.ValidateRecipients())
                {
                    forwardMail.Send();
                    Console.WriteLine("Forwarded the mail(s)");
                }
                else
                {
                    Console.WriteLine("Incorrect recipient(s), mail(s) cannot be forwarded");
                }
            }

            Console.WriteLine("Press any key to exit");
            Console.Read();

            inbox.UnRegisterItemAddEventHandler();
        }


        /// <summary>
        /// This method gets called when an event is triggered by the monitored folder.
        /// </summary>
        /// <param name="Item">Item corresponding to the event</param>
        public static void FolderEvent(IMailItem mail)
        {
            Console.WriteLine("\nEvent: New item from {0} with subject \"{1}\"!!\n", mail.SenderName, mail.Subject);
        }
    }
}
