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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// declare use of the below to get Outllok wrapper classes (a.k.a. Mail*)
using MailSim.OL;

namespace MailSim
{
    class MailSimTest
    {
        /// <summary>
        /// Test module, focusing on Outlook (OOM) wrapper classes (Mail*)
        /// Also serves as an example for Mail* classes usage 
        /// </summary>
        /// <param name="args">Command line argument. It is expected that first argument is aways "/t"</param>
        public void Execute(string[] args)
        {
            // Open connection to Outlook with default profile. Will start Outlook if it is not running
            MailConnection olConnection = new MailConnection();
            MailStore mailStore = null;
            string mailboxName;
            if (args.Length > 1)
            {
                // We will use mailbox with display name specified in arg[1]
                mailboxName = args[1].ToLower();

                // Get all mailboxes (stores) in the profile. 
                //Returns only email stores (skips Public Folders, Delegates, Archives, PSTs)
                MailStore[] stores = olConnection.GetAllMailStores();

                //Search for specific mailbox (store) to use
                foreach (MailStore store in stores)
                {
                    if (store.DisplayName.ToLower() == mailboxName)
                    {
                        mailStore = store;
                    }
                }
                if (mailStore == null)
                {
                    Console.WriteLine("Cannot find store (mailbox) {0} in default profile", mailboxName);
                    return;
                }
            }
            else
            // use default mailbox
            {
                mailStore = olConnection.GetDefaultStore();
                mailboxName = mailStore.DisplayName;
            }

            // Display all top folders in the mailbox
            MailFolder rootFolder = mailStore.GetRootFolder();
            MailFolders rootFolders = rootFolder.GetSubFolders();
            foreach (MailFolder mailFolder in rootFolders)
            {
                Console.WriteLine(mailFolder.FolderPath);
            }

            // Open Inbox and loop through it's top folders
            MailFolder inbox = mailStore.GetDefaultFolder("olFolderInbox");

            // Registering callback for Inbox "ItemAdd" event
            inbox.RegisterItemAddEventHandler(FolderEvent);
            
            MailItems inboxItems = inbox.GetMailItems();
            Console.WriteLine("Inbox has {0} mail items", inbox.MailItemsCount);
            MailFolders inboxSubFolders = inbox.GetSubFolders();
            MailFolder testFolder = null;
            const string testFolderName = "MailSim Test Folder";
            Console.WriteLine("Exploring Inbox Folders");
            if (null != inboxSubFolders)
            {
                foreach (MailFolder mailFolder in inboxSubFolders)
                {
                    if( testFolderName == mailFolder.Name )
                    {
                        testFolder = mailFolder;
                    }
                    Console.WriteLine(mailFolder.FolderPath);
                }
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
                testFolder.GetSubFolders().GetFirst().Delete();
            }

            // Deleting email in inbox
            if (inbox.MailItemsCount > 2)
            {
                MailItem mailToDelete = inbox.GetMailItems().GetFirst();
                Console.WriteLine("Deleting email {0}", mailToDelete.Subject);
                mailToDelete.Delete();
            }

            // Moving email 
            if (inbox.MailItemsCount > 2)
            {
                inbox.GetMailItems().GetFirst().Move(testFolder);
            }
            
            // Sending new email with message attachment to matching users
            Console.WriteLine("Sending new email to matching GAL users");
            MailItem newMail = mailStore.NewMailItem();
            newMail.Subject = "Test Mail from MailSim to GAL users " + (inbox.MailItemsCount + 1).ToString() + " - " + DateTime.Now.TimeOfDay.ToString();
            newMail.Body = "Test from MailSim to matching users";
            if (inbox.GetMailItems().Count > 0)
            {
                MailItem message = inbox.GetMailItems().GetFirst();
                newMail.AddAttachment(message);
            }
            newMail.AddRecipient(mailboxName);

            OLAddressList gal = mailStore.GetGlobalAddressList();
            if (gal != null)
            {
                List<string> users = gal.GetUsers("Mailsim");
                foreach (string userAddress in users)
                {
                    newMail.AddRecipient(userAddress);
                }
            }

            if (newMail.ValidateRecipients)
            {
                newMail.Send();
                Console.WriteLine("Mail to matching users sent!");
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
                List<string> members = gal.GetDLMembers("Mailsim Users");
                if (members != null)
                {
                    foreach (string userAddress in members)
                    {
                        newMail.AddRecipient(userAddress);
                    }
                }
                else
                {
                    Console.WriteLine("ERROR: DL not found");
                }
            }

            if (newMail.ValidateRecipients)
            {
                newMail.Send();
                Console.WriteLine("Mail to DL members sent!");
            }
            else
            {
                Console.WriteLine("Incorrect recipient(s), mail not sent");
            }


            MailItems inboxMailItems = inbox.GetMailItems();
            // Reply All
            if (inbox.MailItemsCount >= 1)
            {
                MailItem replyMail = inboxMailItems.GetFirst().Reply(true);
                // string body = replyMail.Body;
                replyMail.Body = "Reply All by MailSim" + replyMail.Body;
                Console.WriteLine("Message Body:");
                Console.WriteLine(replyMail.Body);
                replyMail.Send();
                Console.WriteLine("Reply All mail sent!");
            }

            // Forward
            if (inbox.MailItemsCount >= 2)
            {
                MailItem forwardMail = inboxMailItems.GetNext().Forward();
                // string body = forwardMail.Body;
                forwardMail.Body = "Forward by MailSim" + forwardMail.Body;
                Console.WriteLine("Message Body:");
                Console.WriteLine(forwardMail.Body);
                forwardMail.AddRecipient(mailboxName);
                if (forwardMail.ValidateRecipients)
                {
                    forwardMail.Send();
                    Console.WriteLine("Forward mail sent!");
                }
                else
                {
                    Console.WriteLine("Incorrect recipient(s), forward mail not sent");
                }
            }

            Console.WriteLine("Hit any key to exit");
            Console.Read();
            inbox.UnRegisterItemAddEventHandler();
            return;
        }


        /// <summary>
        /// This method gets called when an event is triggered by the monitored folder.
        /// </summary>
        /// <param name="Item">Item corresponding to the event</param>
        public static void FolderEvent(MailItem mail)
        {
            Console.WriteLine("\nEvent: New item from {0} with subject \"{1}\"!!\n", mail.SenderName, mail.Subject);
        }
    }
}
