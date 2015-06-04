//--------------------------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized
// to use this sample source code. For the terms of the license, please see the
// license agreement between you and Microsoft.
//--------------------------------------------------------------------------------------
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.Win32;


using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailSim
{
    class OutlookApp
    {
        public Outlook.Application outlook;
        public Dictionary<string, Outlook.OlDefaultFolders> outlookFolder = new Dictionary<string, Outlook.OlDefaultFolders>();
        public List<Outlook.MAPIFolder> eventList = new List<Outlook.MAPIFolder>();

        public const string DefaultMailSubject = "Default MailSim Subject";
        public const string SendMailAction = "sendmail";
        public const string DeleteMailAction = "deletemail";
        public const string ReplyMailAction = "replymail";
        public const string FolderAction = "folder";
        public const string EventMonitorAction = "eventmonitor";

        public const string Sleep = "Sleep";
        public const string Iteration = "Iteration";




        /// <summary>
        /// OutlookApp constructor
        /// </summary>
        public OutlookApp()
        {
            outlook = null;

            // Initalize the supported default Outlook folders
            outlookFolder.Add("olFolderInbox", Outlook.OlDefaultFolders.olFolderInbox);
            outlookFolder.Add("olFolderDeletedItems", Outlook.OlDefaultFolders.olFolderDeletedItems);
            outlookFolder.Add("olFolderDrafts", Outlook.OlDefaultFolders.olFolderDrafts);
            outlookFolder.Add("olFolderJunk", Outlook.OlDefaultFolders.olFolderJunk);
            outlookFolder.Add("olFolderOutbox", Outlook.OlDefaultFolders.olFolderOutbox);
            outlookFolder.Add("olFolderSentMail", Outlook.OlDefaultFolders.olFolderSentMail);
            outlookFolder.Add("olFolderTasks", Outlook.OlDefaultFolders.olFolderTasks);
            outlookFolder.Add("olFolderToDo", Outlook.OlDefaultFolders.olFolderToDo);
            outlookFolder.Add("olPublicFoldersAllPublicFolders", Outlook.OlDefaultFolders.olPublicFoldersAllPublicFolders);

            // Disables the Outlook prompts
//            ConfigOutlookPrompts(false);
        }


        /// <summary>
        /// OutlookApp destructor
        /// </summary>
        ~OutlookApp()
        {
            // Unregisters all events during exit
            for (int count = 0; count < eventList.Count; count ++)
            {
                Console.WriteLine("Cleanup: Removing registered event to folder " + eventList[count].FolderPath);
                RegisterFolderEvent(eventList[count], false);
            }

            // Closes the Outlook process
            if (outlook != null)
                outlook.Quit();

            // Re-enable the Outlook Prompts
//            ConfigOutlookPrompts(true);
        }


        /// <summary>
        /// Determines the action to take
        /// </summary>
        /// <param name="actionNode">XML element node of the operation</param>
        public void ProcessTask(XmlNode actionNode)
        {
            int iteration = 1;
            if (actionNode.Attributes[Iteration] != null)
            {
                iteration = Convert.ToInt32(actionNode.Attributes[Iteration].Value);
            }

            // Sleep after the task is started, if specified
            int sleep = 0;
            if (actionNode.Attributes[Sleep] != null)
            {
                sleep = Convert.ToInt32(actionNode.Attributes[Sleep].Value);
            }

            while (iteration > 0)
            {
                switch (actionNode.Name.ToLower())
                {
                    case SendMailAction:
                        SendMail(actionNode, iteration);
                        break;
                    case DeleteMailAction:
                        DeleteMail(actionNode, iteration);
                        break;
                    case ReplyMailAction:
                        ReplyMail(actionNode);
                        break;
                    case FolderAction:
                        switch (actionNode.Attributes["Task"].Value.ToLower())
                        {
                            case "create":
                                Folder(actionNode, true, iteration);
                                break;
                            case "delete":
                                Folder(actionNode, false, iteration);
                                break;
                            default:
                                Console.WriteLine("Folder: skipping unsupported task " + actionNode.Attributes["Task"].Value);
                                break;
                        }
                        break;
                    case EventMonitorAction:
                        // There is no need to register the same folder event multiple times
                        if (iteration > 1)
                        {
                            Console.WriteLine("EventMonitor: skipping multiple registration of the same event");
                            break;
                        }

                        switch (actionNode.Attributes["Event"].Value.ToLower())
                        {
                            case "register":
                                FolderEventMonitor(actionNode, true);
                                break;
                            case "unregister":
                                FolderEventMonitor(actionNode, false);
                                break;
                            default:
                                Console.WriteLine("EventMonitor: skipping unsupported task " + actionNode.Attributes["Event"].Value);
                                break;
                        }
                        break;

                    default:
                        Console.WriteLine("Skipping unknown action {0} in test config xml", actionNode.Name);
                        break;
                }

                if (sleep != 0)
                {
                    Console.WriteLine("Sleeping for {0} seconds", sleep);
                    Thread.Sleep(sleep * 1000);
                }

                iteration--;
            }
        }

       
        /// <summary>
        /// This method opens and logon to Outlook with the specified profile
        /// </summary>
        /// <param name="profile">Outlook profile to use</param>
        /// <param name="password">Password for the profile</param>
        /// <returns>True if logon is successful, False if logon is failed</returns>
        public bool Logon(string profile, string password)
        {
            try
            {
                // Checks whether an Outlook process is currently running
                if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
                {
                    Console.WriteLine("Outlook is currently running!");
                    return false;

                    /*
                    // If so, use the GetActiveObject method to obtain the process and cast it to an Application object, and close the application
                    outlook = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
                    outlook.Quit();
                    Thread.Sleep(1000);
                     */
                }

                // Create a new instance of Outlook and log on to the specified profile.
                Console.WriteLine("Starting new Outlook session");

                outlook = new Outlook.Application();
                Outlook.NameSpace nameSpace = outlook.GetNamespace("MAPI");
                nameSpace.Logon(profile, password, Missing.Value, Missing.Value);
                nameSpace = null;

                // Wait until the process is running
                int count = 50;
                while (Process.GetProcessesByName("OUTLOOK").Count() <= 0 && count > 0)
                {
                    Thread.Sleep(1000);
                    count--;
                }

                if (count == 0)
                {
                    Console.WriteLine("Error: Unable to start Outlook");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Logon: Exception encountered\n" + ex.ToString());
                return false;
            } 

            return true;
        }


        /// <summary>
        /// This method sends email according to the configuration XML
        /// </summary>
        /// <param name="node">XML element node of the operation</param>
        /// <param name="iteration">Number of iterations to execute this operation</param>
        public void SendMail(XmlNode node, int iteration)
        {
            try
            {
                string subject = DefaultMailSubject;
                string body = "Default MailSim Body";

                // we support either random recipient or specific recipients
                XmlNodeList recipients = node.SelectNodes("Recipient");
                XmlNode randomRecipient = node.SelectSingleNode("RandomRecipient");
                if (recipients.Count == 0)
                {
                    if (randomRecipient == null)
                    {
                        Console.WriteLine("SendMail: error, there is no recipient for sending email");
                        return;
                    }
                    else
                    {
                        // randomly pick recipients from the Address Book
                    }
                }
                else
                {
                    if (randomRecipient != null)
                    {
                        Console.WriteLine("SendMail: error, both RandomRecipient and Recipient exist for SendMail");
                        return;
                    }

                }
                
                XmlNodeList attachments = node.SelectNodes("Attachment");

                if (node.SelectSingleNode("Subject") != null)
                {
                    if (iteration != 1)
                    {
                        subject = node.SelectSingleNode("Subject").InnerText.Replace('"', ' ').Trim() + iteration.ToString();
                    }
                    else
                    {
                        subject = node.SelectSingleNode("Subject").InnerText.Replace('"', ' ').Trim();
                    }
                }

                if (node.SelectSingleNode("Body") != null)
                {
                    body = node.SelectSingleNode("Body").InnerText.Replace('"', ' ').Trim();
                }

                Console.WriteLine("SendMail: Sending mail to {0} recipient(s) with subject {1} and {2} attachment(s)", recipients.Count, subject, attachments.Count);

                Outlook.MailItem mail = outlook.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                mail.Subject = subject;
                mail.Body = body;

                // Add recipient(s)
                foreach(XmlNode xn in recipients)
                {
                    mail.Recipients.Add(xn.InnerText.Replace('"', ' ').Trim());
                }

                if(!mail.Recipients.ResolveAll())
                {
                    Console.WriteLine("SendMail: Unable to resolve recipient name");
                    return;
                }

                // Add attachment(s) if specified
                foreach(XmlNode xn in attachments)
                {
                    if (File.Exists(xn.InnerText))
                    {
                        mail.Attachments.Add(xn.InnerText, Outlook.OlAttachmentType.olByValue,
                            Type.Missing, Type.Missing);
                    }
                    else
                    {
                        Console.WriteLine("SendMail: Skipping non-exist attachment file at " + xn.InnerText);
                    }
                }

                mail.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine("SendMail: Exception encountered\n" + ex.ToString());
            }
        }


        /// <summary>
        /// This method deletes email(s) according to the configuration XML
        /// For a list of supported folders, see http://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.outlook.oldefaultfolders%28v=office.15%29.aspx
        /// </summary>
        /// <param name="node">XML element node of the operation</param>
        /// <param name="iteration">Number of iterations to execute this operation</param>
        public void DeleteMail(XmlNode node, int iteration)
        {
            try
            {
                string subject = DefaultMailSubject;

                if (node.SelectSingleNode("Subject") != null)
                {
                    if (iteration != 1)
                    {
                        subject = node.SelectSingleNode("Subject").InnerText + iteration.ToString();
                    }
                    else
                    {
                        subject = node.SelectSingleNode("Subject").InnerText;
                    }
                }

                Outlook.MAPIFolder folder = GetFolder(node, "Folder");
                if (folder == null)
                {
                    Console.WriteLine("DeleteMail: unable to retrieve folder");
                    return;
                }
                
                Console.WriteLine("DeleteMail: Found {0} items in {1}", folder.Items.Count, folder.Name);

                // List of emails to delete
                List<Outlook.MailItem> emails = new List<Outlook.MailItem>();

                foreach (object item in folder.Items)
                {
                    // NOTE: this is only deleting MailItem right now
                    if (item is Outlook.MailItem && ((Outlook.MailItem)item).Subject == subject)
                    {
                        emails.Add((Outlook.MailItem)item);
                    }
                }

                foreach(Outlook.MailItem item in emails)
                {
                    Console.WriteLine("DeleteMail: Found eamil with subject \"{0}\", deleting it", item.Subject);
                    item.Delete();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("DeleteMail: Exception encountered\n" + ex.ToString());
            }
        }


        /// <summary>
        /// This method replies email according to the configuration XML
        /// </summary>
        /// <param name="node">XML element node of the operation</param>
        public void ReplyMail(XmlNode node)
        {
            try
            {
                // Default mail subject and body
                string replySubject = DefaultMailSubject;
                string replyBody = "";
                bool replyAll = false;

                if (node.SelectSingleNode("MailSubjectToReply") != null)
                {
                    replySubject = node.SelectSingleNode("MailSubjectToReply").InnerText;
                }

                if (node.SelectSingleNode("ReplyBody") != null)
                {
                    replyBody = node.SelectSingleNode("ReplyBody").InnerText;
                }

                if (node.Attributes["ReplyAll"] != null)
                {
                    switch (node.Attributes["ReplyAll"].Value.ToLower())
                    {
                        case "true":
                            replyAll = true;
                            break;
                        case "false":
                            replyAll = false;
                            break;
                        default:
                            Console.WriteLine("ReplyMail: unsupported ReplyAll value " + node.Attributes["ReplyAll"].Value);
                            return;
                    }
                }

                // Retrieves the Outlook folder
                Outlook.MAPIFolder folder = GetFolder(node, "Folder");
                if (folder == null)
                {
                    Console.WriteLine("ReplyMail: unable to retrieve folder");
                    return;
                }

                Console.WriteLine("ReplyMail: Found {0} items in {1}", folder.Items.Count, folder.Name);

                // List of emails to reply
                List<Outlook.MailItem> emails = new List<Outlook.MailItem>();

                foreach (object item in folder.Items)
                {
                    // NOTE: only replyig MailItem right now
                    if (item is Outlook.MailItem && ((Outlook.MailItem)item).Subject == replySubject)
                    {
                        emails.Add((Outlook.MailItem)item);
                    }
                }

                if (emails.Count == 0)
                {
                    Console.WriteLine("ReplyMail: unable to find subject {0} from folder {1}", replySubject, folder.ToString());
                    return;
                }

                Outlook.MailItem reply; 
                foreach (Outlook.MailItem item in emails)
                {
                    if (replyAll)
                    {
                        Console.WriteLine("ReplyMail: Found eamil with subject \"{0}\", replying all", item.Subject);
                        /*
                        Outlook.Action action = item.Actions["Reply"];
                        action.ReplyStyle = Outlook.OlActionReplyStyle.olIncludeOriginalText;
                        reply = action.Execute();
                        */
                        reply = item.ReplyAll();
                    }
                    else
                    {
                        Console.WriteLine("ReplyMail: Found eamil with subject \"{0}\", replying to sender", item.Subject);
                        reply = item.Reply();
                    }
                    reply.Body = replyBody;
                    reply.Send();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ReplyMail: Exception encountered\n" + ex.ToString());
            }
        }


        public void ForwardMail(XmlNode node)
        {
            try
            {
                // Default mail subject and body
                string replySubject = DefaultMailSubject;
                string replyBody = "";
                bool replyAll = false;

                if (node.SelectSingleNode("MailSubjectToReply") != null)
                {
                    replySubject = node.SelectSingleNode("MailSubjectToReply").InnerText;
                }

                if (node.SelectSingleNode("ReplyBody") != null)
                {
                    replyBody = node.SelectSingleNode("ReplyBody").InnerText;
                }

                if (node.Attributes["ReplyAll"] != null)
                {
                    switch (node.Attributes["ReplyAll"].Value.ToLower())
                    {
                        case "true":
                            replyAll = true;
                            break;
                        case "false":
                            replyAll = false;
                            break;
                        default:
                            Console.WriteLine("ReplyMail: unsupported ReplyAll value " + node.Attributes["ReplyAll"].Value);
                            return;
                    }
                }

                // Retrieves the Outlook folder
                Outlook.MAPIFolder folder = GetFolder(node, "Folder");
                if (folder == null)
                {
                    Console.WriteLine("ReplyMail: unable to retrieve folder");
                    return;
                }

                Console.WriteLine("ReplyMail: Found {0} items in {1}", folder.Items.Count, folder.Name);

                // List of emails to reply
                List<Outlook.MailItem> emails = new List<Outlook.MailItem>();

                foreach (object item in folder.Items)
                {
                    // NOTE: only replyig MailItem right now
                    if (item is Outlook.MailItem && ((Outlook.MailItem)item).Subject == replySubject)
                    {
                        emails.Add((Outlook.MailItem)item);
                    }
                }

                if (emails.Count == 0)
                {
                    Console.WriteLine("ReplyMail: unable to find subject {0} from folder {1}", replySubject, folder.ToString());
                    return;
                }

                Outlook.MailItem reply;
                foreach (Outlook.MailItem item in emails)
                {
                    if (replyAll)
                    {
                        Console.WriteLine("ReplyMail: Found eamil with subject \"{0}\", replying all", item.Subject);
                        /*
                        Outlook.Action action = item.Actions["Reply"];
                        action.ReplyStyle = Outlook.OlActionReplyStyle.olIncludeOriginalText;
                        reply = action.Execute();
                        */
                        reply = item.ReplyAll();
                    }
                    else
                    {
                        Console.WriteLine("ReplyMail: Found eamil with subject \"{0}\", replying to sender", item.Subject);
                        reply = item.Reply();
                    }
                    reply.Body = replyBody;
                    reply.Send();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ReplyMail: Exception encountered\n" + ex.ToString());
            }
        }


        public void MailCopy(XmlNode node)
        {
            /*
private void items_ItemAdd(object Item)
 {


    Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)
    this.Application.ActiveExplorer().Session.GetDefaultFolder
    (Outlook.OlDefaultFolders.olFolderInbox);


    // the incoming email
    Outlook.MailItem mail = (Outlook.MailItem)Item;
    //make a copy of it but error occurs
    Outlook.MailItem cItem = mail.copy();
    //
    cItem = (Outlook.MailItem)cItem.Move((Outlook.MAPIFolder)
    this.Application.ActiveExplorer().Session.GetDefaultFolder
    (Outlook.OlDefaultFolders.olFolderJunk));

             */
        }

        /// <summary>
        /// This method creates or deletes folder from Outlook according to the configuration XML
        /// </summary>
        /// <param name="node">XML element node of the operation</param>
        /// <param name="create">True to create a new folder. False to delete a folder</param>
        /// <param name="iteration">Number of iterations to execute this operation</param>
        public void Folder(XmlNode node, bool create, int iteration)
        {
            try
            {
                string folderName = "";

                if (node.SelectSingleNode("FolderName") != null)
                {
                    // Outlook requires a unique folder name for each folder
                    if (iteration != 1)
                    {
                        folderName = node.SelectSingleNode("FolderName").InnerText + " " + iteration.ToString();
                    }
                    else
                    {
                        folderName = node.SelectSingleNode("FolderName").InnerText;
                    }
                }
                else
                {
                    Console.WriteLine("Folder: Folder element is not specified in the config file");
                    return;
                }

                // Retrieves the Outlook folder
                Outlook.MAPIFolder folder = GetFolder(node, "FolderPath");
                if (folder == null)
                {
                    Console.WriteLine("Folder: unable to retrieve folder");
                    return;
                }

                Console.WriteLine("Folder: Found {0} folders in {1}", folder.Folders.Count, folder.Name);

                if (create)
                {
                    foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                    {
                        if (subFolder.Name == folderName)
                        {
                            Console.WriteLine("Folder: folder {0} already exists, skipping new folder", folderName);
                            return;
                        }
                    }

                    Outlook.MAPIFolder newfolder = folder.Folders.Add(folderName);
                    if (newfolder != null)
                    {
                        Console.WriteLine("Folder: Created new folder {0} at {1}", newfolder.Name, newfolder.FolderPath);
                    }
                }
                else
                {
                    bool deleted = false;
                    foreach (Outlook.MAPIFolder subFolder in folder.Folders)
                    {
                        if (subFolder.Name == folderName)
                        {
                            Console.WriteLine("Folder: deleting folder {0}", subFolder.FolderPath);
                            subFolder.Delete();
                            deleted = true;
                        }
                    }
                    if (!deleted)
                    {
                        Console.WriteLine("Folder: unable to find folder {0} to delete", folderName);
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Folder: Exception encountered\n" + ex.ToString());
            }
        }


        /// <summary>
        /// This method handles event registration
        /// </summary>
        /// <param name="node">XML element node of the operation</param>
        /// <param name="register">True to register folder event monitoring, False to unregister folder event monitoring</param>
        public void FolderEventMonitor(XmlNode node, bool register)
        {
            try
            {
                Outlook.MAPIFolder folder = GetFolder(node, "Folder");
                if (folder == null)
                {
                    Console.WriteLine("EventMonitor: unable to retrieve folder");
                    return;
                }
                RegisterFolderEvent(folder, register);
            }
            catch (Exception ex)
            {
                Console.WriteLine("RegisterEvent: Exception encountered\n" + ex.ToString());
            }
        }


        /// <summary>
        /// This method registers or unregisters folder event monitoring.
        /// </summary>
        /// <param name="folder">Outlook folder</param>
        /// <param name="register">True to register folder event monitoring, False to unregister folder event monitoring</param>
        public void RegisterFolderEvent(Outlook.MAPIFolder folder, bool register)
        {
            try
            {
                if (register)
                {
                    folder.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(FolderEvent);
                    eventList.Add(folder);
                    Console.WriteLine("EventMonitor: Registered event to " + folder.FolderPath);
                }
                else
                {
                    folder.Items.ItemAdd -= new Outlook.ItemsEvents_ItemAddEventHandler(FolderEvent);
                    RemoveFolderEventList(folder);
                    Console.WriteLine("EventMonitor: Unregistered event to " + folder.FolderPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("RegisterEvent: Exception encountered\n" + ex.ToString());
            }
        }


        /// <summary>
        /// This method finds and removes registered event from the saved list.
        /// </summary>
        /// <param name="folder">Outlook folder where the event was registered</param>
        private void RemoveFolderEventList(Outlook.MAPIFolder folder)
        {
            if (eventList.Count == 0)
            {
                Console.WriteLine("RemoveEventList: Monitor event list is empty and skipping " + folder.FolderPath);
                return;
            }

            // Uses the folder path to determine if the events are the same
            for (int count = 0; count < eventList.Count; count++)
            {
                if (eventList[count].FolderPath == folder.FolderPath)
                {
                    eventList.RemoveAt(count);
                    return;
                }
            }
        }
        

        /// <summary>
        /// This method gets called when an event is triggered by the monitored folder.
        /// </summary>
        /// <param name="Item">Item corresponding to the event</param>
        public static void FolderEvent(object Item)
        {
            // Supports only the MailItem
            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem mail = (Outlook.MailItem)Item;
                if (Item != null)
                {
                    Console.WriteLine("\nEvent: New item from {0} with subject \"{1}\"!!\n", mail.Sender.Name, mail.Subject);
                }
            }
            else
            {
                Console.WriteLine("Event received but with unknown type " + Item.GetType().ToString());
            }
        }


        /// <summary>
        /// This method retreives the folder element from the configuration XML node and converts 
        /// it to the Outlook MAPIFolder property.
        /// This method only supports Outlook default folder.
        /// List of supported folders can be found in http://msdn.microsoft.com/en-us/library/office/microsoft.office.interop.outlook.oldefaultfolders%28v=office.15%29.aspx
        /// </summary>
        /// <param name="node">XML element node of the operation</param>
        /// <param name="folder">Name of the node element that represents the folder path</param>
        /// <returns>Outlook MAPIFolder that corresponding to the specified folder</returns>
        private Outlook.MAPIFolder GetFolder(XmlNode node, string folder)
        {
            string folderStr = ""; 

            if (node == null || folder == "")
            {
                Console.WriteLine("GetFolder: node is null or folder is unspecified");
                return null;
            }

            if (node.SelectSingleNode(folder) != null)
            {
                folderStr = node.SelectSingleNode(folder).InnerText;
            }
            else
            {
                Console.WriteLine("GetFolder: Folder element is not specified in the config file");
            }

            // Converts the folder string to the supported folder type
            if (!outlookFolder.ContainsKey(folderStr))
            {
                Console.WriteLine("GetFolder: Unsupported folder {0}", folderStr);
                Console.Write("Supported folder: ");
                foreach (KeyValuePair<string, Outlook.OlDefaultFolders> pair in outlookFolder)
                {
                    Console.Write("\t{0}\n", pair.Value);
                }
                return null;
            }

            // We are only supporting the default folder right now
            return outlook.GetNamespace("MAPI").GetDefaultFolder(outlookFolder[folderStr]);
        }



    }
}


