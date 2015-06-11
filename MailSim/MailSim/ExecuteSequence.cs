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
using System.Threading;
using System.Xml;
using Microsoft.Win32;

using MailSim.OL;

namespace MailSim
{
    class ExecuteSequence
    {
        private MailSimSequence sequence;
        private MailSimOperations operations;
        private XmlDocument operationXML;
        private MailConnection olConnection;
        private MailStore olMailStore;
        private Random randomNum;

        private const string OfficeVersion = "15.0";
        private const string OfficePolicyRegistryRoot = @"Software\Policies\Microsoft\Office\" + OfficeVersion;
        private const string OutlookPolicyRegistryRoot = OfficePolicyRegistryRoot + @"\Outlook";

        private const string Recipients = "Recipients";
        private const string RandomRecipients = "RandomRecipients";
        private const string DefaultSubject = "Default Subject";
        private const string DefaultBody = "Default Body";
        private const int MaxNumberOfRandomFolder = 100;

        private List<MailFolder> FolderEventList = new List<MailFolder>();


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="seq">Sequence file content </param>
        public ExecuteSequence(MailSimSequence seq)
        {
            if (seq != null)
            {
                try
                {
                    sequence = seq;

                    // disables the Outlook security prompt if specified
                    if (sequence.DisableOutlookPrompt == true)
                    {
                        ConfigOutlookPrompts(false);
                    }

                    // Openes connection to Outlook with default profile, starts Outlook if it is not running
                    olConnection = new MailConnection();

                    // note: we are only supporting the default Mail Store right now
                    olMailStore = olConnection.GetDefaultStore();

                    // initializes a random number
                    randomNum = new Random();
                }
                catch (Exception)
                {
                    Log.Out(Log.Severity.Error, "Execute", "Error encountered during initialization");
                    throw;
                }
            }
        }


        /// <summary>
        /// Destructor
        /// </summary>
        ~ExecuteSequence()
        {
            if (sequence == null)
                return;

            // restore the Outlook prompt if needed
            if (sequence.DisableOutlookPrompt == true)
            {
                ConfigOutlookPrompts(true);
            }
        }


        /// <summary>
        /// This method unregisters event 
        /// </summary>
        public void CleanupAfterIteration()
        {
            // unregisters all registered folder event 
            foreach (MailFolder folder in FolderEventList)
            {
                RegisterFolderEvent("Event", folder, false);
            }

            FolderEventList.Clear();

        }


        /// <summary>
        /// This method process the sequence file
        /// </summary>
        public void Execute()
        {
            if (sequence == null)
            {
                Log.Out(Log.Severity.Error, "Execute", "Sequence content is empty!");
                return;
            }

            // registers to monitor the Inbox
            MailSimOperationsEventMonitor inboxEvent = new MailSimOperationsEventMonitor();
            inboxEvent.Folder = "olFolderInbox";
            inboxEvent.OperationName = "DefaultInboxMonitor";
            EventMonitor(inboxEvent);

            // process each operation group
            foreach (MailSimSequenceOperationGroup group in sequence.OperationGroup)
            {
                int iterations = 1;
                if (!string.IsNullOrEmpty(group.Iterations))
                {
                    iterations = Convert.ToInt32(group.Iterations);
                }

                // processes the operations file
                operations = ConfigurationFile.LoadOperationFile(group.OperationFile, out operationXML);

                if (operations == null)
                {
                    Log.Out(Log.Severity.Error, group.Name, "Skipping OperationGroup");
                    continue;
                }

                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, group.Name, "Starting group run {0}", count);

                    foreach (MailSimSequenceOperationGroupTask task in group.Task)
                    {
                        ProcessTask(task);
                    }

                    Log.Out(Log.Severity.Info, group.Name, "Finished group run {0}", count);

                    if (!string.IsNullOrEmpty(group.Sleep))
                    {
                        int sleep = Convert.ToInt32(group.Sleep);
                        Log.Out(Log.Severity.Info, group.Name, "Sleeping for {0} seconds", sleep);
                        Thread.Sleep(sleep * 1000);
                    }
                }

                CleanupAfterIteration();
            }

            return;
        }


        /// <summary>
        /// This method processes each task of the OperationGroup, hanlding the iteration and sleep
        /// </summary>
        /// <param name="task">task to process</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        public void ProcessTask(MailSimSequenceOperationGroupTask task)
        {
            int iterations = 1;

            if (!string.IsNullOrEmpty(task.Iterations))
            {
                iterations = Convert.ToInt32(task.Iterations);
            }

            for (int count = 1; count <= iterations; count++)
            {
                Log.Out(Log.Severity.Info, task.Name, "Processing task run {0}", count);

                if (ExecuteTask(task.Name))
                {
                    Log.Out(Log.Severity.Info, task.Name, "Finished processing task run {0}", count);
                }
                else
                {
                    Log.Out(Log.Severity.Error, task.Name, "Failed processing task");
                }

                if (!string.IsNullOrEmpty(task.Sleep))
                {
                    int sleep = Convert.ToInt32(task.Sleep);
                    Log.Out(Log.Severity.Info, task.Name, "Sleeping for {0} seconds", sleep);
                    Thread.Sleep(sleep * 1000);
                }
            }
        }


        /// <summary>
        /// This method determines and calls the appropiate method to execute the task
        /// </summary>
        /// <param name="taskName">name of the task</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        public bool ExecuteTask(string taskName)
        {
            // locates the actual operation using the name
            XmlNodeList opNodes = operationXML.SelectNodes("//MailSimOperations/*[@OperationName='" + taskName + "']");

            // we expect only 1 operation node matching the name
            if (opNodes.Count != 1)
            {
                Log.Out(Log.Severity.Error, taskName,
                    "There are {0} nodes in the operation XML file with name {1}",
                    opNodes.Count, taskName);
                return false;
            }

            foreach (object operation in operations.Items)
            {
                // determine the right operation for the task
                if (operation.GetType() == typeof(MailSimOperationsMailSend))
                {
                    if (((MailSimOperationsMailSend)operation).OperationName == taskName)
                    {
                        return MailSend((MailSimOperationsMailSend)operation);
                    }
                }
                else if (operation.GetType() == typeof(MailSimOperationsMailDelete))
                {
                    if (((MailSimOperationsMailDelete)operation).OperationName == taskName)
                    {
                        return MailDelete((MailSimOperationsMailDelete)operation);
                    }
                }
                else if (operation.GetType() == typeof(MailSimOperationsMailReply))
                {
                    if (((MailSimOperationsMailReply)operation).OperationName == taskName)
                    {
                        return MailReply((MailSimOperationsMailReply)operation);
                    }
                }
                else if (operation.GetType() == typeof(MailSimOperationsMailForward))
                {
                    if (((MailSimOperationsMailForward)operation).OperationName == taskName)
                    {
                        return MailForward((MailSimOperationsMailForward)operation);
                    }
                }
                else if (operation.GetType() == typeof(MailSimOperationsMailMove))
                {
                    if (((MailSimOperationsMailMove)operation).OperationName == taskName)
                    {
                        return MailMove((MailSimOperationsMailMove)operation);
                    }
                }
                else if (operation.GetType() == typeof(MailSimOperationsFolderCreate))
                {
                    if (((MailSimOperationsFolderCreate)operation).OperationName == taskName)
                    {
                        return FolderCreate((MailSimOperationsFolderCreate)operation);
                    }
                }
                else if (operation.GetType() == typeof(MailSimOperationsFolderDelete))
                {
                    if (((MailSimOperationsFolderDelete)operation).OperationName == taskName)
                    {
                        return FolderDelete((MailSimOperationsFolderDelete)operation);
                    }
                }
                else if (operation.GetType() == typeof(MailSimOperationsEventMonitor))
                {
                    if (((MailSimOperationsEventMonitor)operation).OperationName == taskName)
                    {
                        return EventMonitor((MailSimOperationsEventMonitor)operation);
                    }
                }
                else
                {
                    Log.Out(Log.Severity.Error, taskName, "Skipping unknown task");
                }
            }

            Log.Out(Log.Severity.Error, taskName, "Unable to find matching task, skipping task");
            return false;
        }


        /// <summary>
        /// This method sends mail according to the paramenter
        /// </summary>
        /// <param name="operation">parameters for MailSend</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailSend(MailSimOperationsMailSend operation)
        {
            int iterations = 1;
            if (!string.IsNullOrEmpty(operation.Count))
            {
                iterations = Convert.ToInt32(operation.Count);
            }

            if (iterations < 1)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Count is less than the minimum allowed value", iterations);
                return false;
            }

            for (int count = 1; count <= iterations; count++)
            {
                List<string> recipients = GetRecipients(operation.OperationName, operation.RecipientType, operation.Recipients);

                if (recipients == null)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Recipient is not specified, skipping operation");
                    return false;
                }

                List<string> attachments = GetAttachments(operation.OperationName, operation.Attachments);

                try
                {
                    // generates a new email
                    MailItem mail = olMailStore.NewMailItem();

                    mail.Subject = mail.Body = System.DateTime.Now.ToString() + " - ";
                    mail.Subject += (string.IsNullOrEmpty(operation.Subject)) ? DefaultSubject : operation.Subject;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Subject: {0}", mail.Subject);
                    mail.Body += (string.IsNullOrEmpty(operation.Body)) ? DefaultBody : operation.Body;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Body: {0}", mail.Body);

                    // adds all recipients
                    foreach (string recpt in recipients)
                    {
                        Log.Out(Log.Severity.Info, operation.OperationName, "Recipient: {0}", recpt);
                        mail.AddRecipient(recpt);
                    }

                    // processes the attachment
                    foreach (string attmt in attachments)
                    {
                        Log.Out(Log.Severity.Info, operation.OperationName, "Attachment: {0}", attmt);
                        mail.AddAttachment(attmt);
                    }

                    mail.Send();
                }
                catch (Exception ex)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Exception encountered\n" + ex);
                    return false;
                }

                if (!string.IsNullOrEmpty(operation.Sleep))
                {
                    int sleep = Convert.ToInt32(operation.Sleep);
                    Log.Out(Log.Severity.Info, operation.OperationName, "Sleeping for {0} seconds", sleep);
                    Thread.Sleep(sleep * 1000);
                }
            }

            return true;
        }


        /// <summary>
        /// This method deletes mail according to the paramenter
        /// </summary>
        /// <param name="operation">parameters for MailDelete</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailDelete(MailSimOperationsMailDelete operation)
        {
            int iterations = 1;
            bool random = false;
            if (!string.IsNullOrEmpty(operation.Count))
            {
                iterations = Convert.ToInt32(operation.Count);
            }

            try
            {
                // retrieves mails from Outlook
                List<MailItem> mails = GetMails(operation.OperationName, operation.Folder, operation.Subject);
                if (mails == null || mails.Count == 0)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Skipping MailDelete");
                    return false;
                }

                // randomly generate the number of emails to delete 
                if (iterations == 0)
                {
                    random = true;
                    iterations = randomNum.Next(1, mails.Count + 1);
                    Log.Out(Log.Severity.Info, operation.OperationName, "Randomly deleting {0} emails", iterations);
                }

                // we need to make sure we are not deleting more than what we have in the mailbox
                if (iterations > mails.Count)
                {
                    Log.Out(Log.Severity.Warning, operation.OperationName,
                        "Only {0} email(s) in the folder, adjusting the number of emails to delete from {0} to {1}",
                        iterations, mails.Count);
                    iterations = mails.Count;
                }

                int indexToDelete;
                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, operation.OperationName, "Starting iteration {0}", count);

                    // just delete the email in order if random is not selected,
                    // otherwise randomly pick the mail to delete
                    indexToDelete = random ? randomNum.Next(0, mails.Count) : mails.Count - 1;

                    Log.Out(Log.Severity.Info, operation.OperationName, "Deleting email with subject: {0}", mails[indexToDelete].Subject);
                    mails[indexToDelete].Delete();
                    mails.RemoveAt(indexToDelete);

                    if (!string.IsNullOrEmpty(operation.Sleep))
                    {
                        int sleep = Convert.ToInt32(operation.Sleep);
                        Log.Out(Log.Severity.Info, operation.OperationName, "Sleeping for {0} seconds", sleep);
                        Thread.Sleep(sleep * 1000);
                    }

                    Log.Out(Log.Severity.Info, operation.OperationName, "Finished iteration {0}", count);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Exception encountered\n" + ex);
                return false;
            }

            return true;
        }


        /// <summary>
        /// This method replies email according to the parameters
        /// </summary>
        /// <param name="operation">parameters for MailReply</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailReply(MailSimOperationsMailReply operation)
        {
            int iterations = 1;
            bool random = false;
            if (!string.IsNullOrEmpty(operation.Count))
            {
                iterations = Convert.ToInt32(operation.Count);
            }

            try
            {
                // retrieves mails from Outlook
                List<MailItem> mails = GetMails(operation.OperationName, operation.Folder, operation.MailSubjectToReply);
                if (mails == null || mails.Count == 0)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Skipping MailReply");
                    return false;
                }

                // randomly generate the number of emails to reply 
                if (iterations == 0)
                {
                    random = true;
                    iterations = randomNum.Next(1, mails.Count + 1);
                    Log.Out(Log.Severity.Info, operation.OperationName, "Randomly replying {0} emails", iterations);
                }

                // we need to make sure we are not replying more than what we have in the mailbox
                if (iterations > mails.Count)
                {
                    Log.Out(Log.Severity.Warning, operation.OperationName,
                        "Only {0} email(s) in the folder, adjusting the number of emails to reply from {0} to {1}",
                        iterations, mails.Count);
                    iterations = mails.Count;
                }

                int indexToReply;
                MailItem mailToReply;
                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, operation.OperationName, "Starting iteration {0}", count);

                    List<string> attachments = GetAttachments(operation.OperationName, operation.Attachments);

                    // just reply the email in order if random is not selected,
                    // otherwise randomly pick the mail to reply
                    indexToReply = random ? randomNum.Next(0, mails.Count) : count - 1;
                    mailToReply = mails[indexToReply].Reply(operation.ReplyAll);

                    Log.Out(Log.Severity.Info, operation.OperationName, "Subject: {0}", mailToReply.Subject);

                    mailToReply.Body = System.DateTime.Now.ToString() + " - " +
                        ((string.IsNullOrEmpty(operation.ReplyBody)) ? DefaultBody : operation.ReplyBody) +
                        mailToReply.Body;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Body: {0}", mailToReply.Body);

                    // process the attachment
                    foreach (string attmt in attachments)
                    {
                        Log.Out(Log.Severity.Info, operation.OperationName, "Attachment: {0}", attmt);
                        mailToReply.AddAttachment(attmt);
                    }

                    mailToReply.Send();

                    if (!string.IsNullOrEmpty(operation.Sleep))
                    {
                        int sleep = Convert.ToInt32(operation.Sleep);
                        Log.Out(Log.Severity.Info, operation.OperationName, "Sleeping for {0} seconds", sleep);
                        Thread.Sleep(sleep * 1000);
                    }

                    Log.Out(Log.Severity.Info, operation.OperationName, "Finished iteration {0}", count);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Exception encountered\n" + ex);
                return false;
            }

            return true;
        }


        /// <summary>
        /// This method forwards emails according to the parameters
        /// </summary>
        /// <param name="operation">arguement for MailForward</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailForward(MailSimOperationsMailForward operation)
        {
            int iterations = 1;
            bool random = false;
            if (!string.IsNullOrEmpty(operation.Count))
            {
                iterations = Convert.ToInt32(operation.Count);
            }

            try
            {
                // retrieves mails from Outlook
                List<MailItem> mails = GetMails(operation.OperationName, operation.Folder, operation.MailSubjectToForward);
                if (mails == null || mails.Count == 0)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Skipping MailForward");
                    return false;
                }

                // randomly generate the number of emails to forward 
                if (iterations == 0)
                {
                    random = true;
                    iterations = randomNum.Next(1, mails.Count + 1);
                    Log.Out(Log.Severity.Info, operation.OperationName, "Randomly forwarding {0} emails", iterations);
                }

                // we need to make sure we are not forwarding more than what we have in the mailbox
                if (iterations > mails.Count)
                {
                    Log.Out(Log.Severity.Warning, operation.OperationName,
                        "Only {0} email(s) in the folder, adjusting the number of emails to forward from {0} to {1}",
                        iterations, mails.Count);
                    iterations = mails.Count;
                }

                int indexToForward;
                MailItem mailToForward;
                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, operation.OperationName, "Starting iteration {0}", count);

                    List<string> recipients = GetRecipients(operation.OperationName, operation.RecipientType, operation.Recipients);

                    if (recipients == null)
                    {
                        Log.Out(Log.Severity.Error, operation.OperationName, "Recipient is not specified, skipping operation");
                        return false;
                    }

                    List<string> attachments = GetAttachments(operation.OperationName, operation.Attachments);

                    // just forward the email in order if random is not selected,
                    // otherwise randomly pick the mail to forward
                    indexToForward = random ? randomNum.Next(0, mails.Count) : count - 1;
                    mailToForward = mails[indexToForward].Forward();

                    Log.Out(Log.Severity.Info, operation.OperationName, "Subject: {0}", mailToForward.Subject);

                    mailToForward.Body = System.DateTime.Now.ToString() + " - " +
                        ((string.IsNullOrEmpty(operation.ForwardBody)) ? DefaultBody : operation.ForwardBody) +
                        mailToForward.Body;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Body: {0}", mailToForward.Body);

                    // adds all recipients
                    foreach (string recpt in recipients)
                    {
                        Log.Out(Log.Severity.Info, operation.OperationName, "Recipient: {0}", recpt);
                        mailToForward.AddRecipient(recpt);
                    }

                    // processes the attachment
                    foreach (string attmt in attachments)
                    {
                        Log.Out(Log.Severity.Info, operation.OperationName, "Attachment: {0}", attmt);
                        mailToForward.AddAttachment(attmt);
                    }

                    mailToForward.Send();

                    if (!string.IsNullOrEmpty(operation.Sleep))
                    {
                        int sleep = Convert.ToInt32(operation.Sleep);
                        Log.Out(Log.Severity.Info, operation.OperationName, "Sleeping for {0} seconds", sleep);
                        Thread.Sleep(sleep * 1000);
                    }

                    Log.Out(Log.Severity.Info, operation.OperationName, "Finished iteration {0}", count);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Exception encountered\n" + ex);
                return false;
            }

            return true;
        }


        /// <summary>
        /// This method moves emails according to the parameters
        /// </summary>
        /// <param name="operation">arguement for MaiMove</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailMove(MailSimOperationsMailMove operation)
        {
            int iterations = 1;
            bool random = false;
            if (!string.IsNullOrEmpty(operation.Count))
            {
                iterations = Convert.ToInt32(operation.Count);
            }

            try
            {
                // retrieves mails from Outlook
                List<MailItem> mails = GetMails(operation.OperationName, operation.SourceFolder, operation.Subject);
                if (mails == null || mails.Count == 0)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Skipping MailMove");
                    return false;
                }

                // randomly generate the number of emails to forward 
                if (iterations == 0)
                {
                    random = true;
                    iterations = randomNum.Next(1, mails.Count + 1);
                    Log.Out(Log.Severity.Info, operation.OperationName, "Randomly moving {0} emails", iterations);
                }

                // we need to make sure we are not moving more than what we have in the mailbox
                if (iterations > mails.Count)
                {
                    Log.Out(Log.Severity.Warning, operation.OperationName,
                        "Only {0} email(s) in the folder, adjusting the number of emails to move from {0} to {1}",
                        iterations, mails.Count);
                    iterations = mails.Count;
                }

                // gets the Outlook destination folder
                MailFolder destinationFolder = olMailStore.GetDefaultFolder(operation.DestinationFolder);
                if (destinationFolder == null)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Unable to retrieve folder {0}",
                        operation.DestinationFolder);
                    return false;
                }

                int indexToCopy;
                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, operation.OperationName, "Starting iteration {0}", count);

                    // just copy the email in order if random is not selected,
                    // otherwise randomly pick the mail to copy
                    indexToCopy = random ? randomNum.Next(0, mails.Count) : mails.Count - 1;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Move to {0}: {1}",
                        operation.DestinationFolder, mails[indexToCopy].Subject);

                    mails[indexToCopy].Move(destinationFolder);

                    if (!string.IsNullOrEmpty(operation.Sleep))
                    {
                        int sleep = Convert.ToInt32(operation.Sleep);
                        Log.Out(Log.Severity.Info, operation.OperationName, "Sleeping for {0} seconds", sleep);
                        Thread.Sleep(sleep * 1000);
                    }

                    Log.Out(Log.Severity.Info, operation.OperationName, "Finished iteration {0}", count);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Exception encountered\n" + ex);
                return false;
            }
            return true;
        }


        /// <summary>
        /// This method creates folders according to the parameters
        /// </summary>
        /// <param name="operation">parameters for FolderCreate</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool FolderCreate(MailSimOperationsFolderCreate operation)
        {
            int iterations = 1;
            if (!string.IsNullOrEmpty(operation.Count))
            {
                iterations = Convert.ToInt32(operation.Count);
            }

            try
            {
                // gets the Outlook folder
                MailFolder folder = olMailStore.GetDefaultFolder(operation.FolderPath);
                if (folder == null)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Unable to retrieve folder {0}",
                        operation.FolderPath);
                    return false;
                }

                // randomly generate the number of folders to create 
                if (iterations == 0)
                {
                    iterations = randomNum.Next(1, MaxNumberOfRandomFolder + 1);
                    Log.Out(Log.Severity.Info, operation.OperationName, "Randomly creating {0} folders", iterations);
                }

                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, operation.OperationName, "Starting iteration {0}", count);
                    string newFolderName = System.DateTime.Now.ToString() + " - " + operation.FolderName;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Creating folder: {0}", newFolderName);
                    folder.AddSubFolder(newFolderName);

                    if (!string.IsNullOrEmpty(operation.Sleep))
                    {
                        int sleep = Convert.ToInt32(operation.Sleep);
                        Log.Out(Log.Severity.Info, operation.OperationName, "Sleeping for {0} seconds", sleep);
                        Thread.Sleep(sleep * 1000);
                    }

                    Log.Out(Log.Severity.Info, operation.OperationName, "Finished iteration {0}", count);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Exception encountered\n" + ex);
                return false;
            }

            return true;
        }


        /// <summary>
        /// This method deletes folders according to the parameters
        /// </summary>
        /// <param name="operation">parameters for FolderDelete</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool FolderDelete(MailSimOperationsFolderDelete operation)
        {
            int iterations = 1;
            bool random = false;
            if (!string.IsNullOrEmpty(operation.Count))
            {
                iterations = Convert.ToInt32(operation.Count);
            }

            try
            {
                // gets the Outlook folder
                MailFolder folder = olMailStore.GetDefaultFolder(operation.FolderPath);
                if (folder == null)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Unable to retrieve folder {0}",
                        operation.FolderPath);
                    return false;
                }

                List<MailFolder> subFolders = GetMatchingSubFolders(operation.OperationName, folder, operation.FolderName);
                if (subFolders.Count == 0)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "There is no matching folder to delete in folder {0}",
                        folder.Name);
                    return false;
                }

                // randomly generate the number of folder to delete if Count is not specified
                if (iterations == 0)
                {
                    random = true;
                    iterations = randomNum.Next(1, subFolders.Count + 1);
                    Log.Out(Log.Severity.Info, operation.OperationName, "Randomly deleting {0} folders", iterations);
                }
                else if (iterations > subFolders.Count)
                {
                    Log.Out(Log.Severity.Warning, operation.OperationName, "Only {0} folders available, adjusting delete from {1} folders to {0}",
                        subFolders.Count, iterations);
                    iterations = subFolders.Count;
                }

                int indexToDelete;
                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, operation.OperationName, "Starting iteration {0}", count);

                    indexToDelete = random ? randomNum.Next(0, subFolders.Count) : subFolders.Count - 1;

                    // deletes the folder and remove it from the saved list
                    Log.Out(Log.Severity.Info, operation.OperationName, "Deleting folder: {0}", subFolders[indexToDelete].Name);
                    subFolders[indexToDelete].Delete();
                    subFolders.RemoveAt(indexToDelete);

                    if (!string.IsNullOrEmpty(operation.Sleep))
                    {
                        int sleep = Convert.ToInt32(operation.Sleep);
                        Log.Out(Log.Severity.Info, operation.OperationName, "Sleeping for {0} seconds", sleep);
                        Thread.Sleep(sleep * 1000);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Exception encountered\n" + ex);
                return false;
            }

            return true;
        }


        /// <summary>
        /// This method registers events to monitor folders
        /// We only support default Outlook folders
        /// </summary>
        /// <param name="operation">parameters for EventMonitor</param>
        /// <returns></returns>
        private bool EventMonitor(MailSimOperationsEventMonitor operation)
        {
            try
            {
                // gets the default Outlook folder
                MailFolder folder = olMailStore.GetDefaultFolder(operation.Folder);
                if (folder == null)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Unable to retrieve folder {0}",
                     operation.Folder);
                    return false;
                }

                // makes sure we are not already registered for the folder
                if (FolderEventList.Contains(folder))
                {
                    Log.Out(Log.Severity.Warning, operation.OperationName, "Event already registered for folder {0}", operation.Folder);
                    return true;
                }
                // registers the event and remember the folder
                else
                {
                    // registers folder event
                    if (RegisterFolderEvent(operation.OperationName, folder, true) != true)
                    {
                        Log.Out(Log.Severity.Error, operation.OperationName, "Unable to register event");
                        return false;
                    }
                }

                if (!string.IsNullOrEmpty(operation.Sleep))
                {
                    int sleep = Convert.ToInt32(operation.Sleep);
                    Log.Out(Log.Severity.Info, operation.OperationName, "Sleeping for {0} seconds", sleep);
                    Thread.Sleep(sleep * 1000);
                }

            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Exception encountered\n" + ex);
                return false;
            }

            return true;
        }


        /// <summary>
        /// This method registers or unregisters folder event monitoring.
        /// </summary>
        /// <param name="operation">name of the operation</param>
        /// <param name="folder">Outlook folder</param>
        /// <param name="register">True to register folder event monitoring, False to unregister folder event monitoring</param>
        private bool RegisterFolderEvent(string operation, MailFolder folder, bool register)
        {
            try
            {
                if (register)
                {
                    folder.RegisterItemAddEventHandler(FolderMonitorEvent);
                    FolderEventList.Add(folder);
                    Log.Out(Log.Severity.Info, operation, "Registered event to " + folder.FolderPath);
                }
                else
                {
                    folder.UnRegisterItemAddEventHandler();
                    Log.Out(Log.Severity.Info, operation, "Unregistered event to " + folder.FolderPath);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, operation, "RegisterFolderEvent: Exception encountered\n" + ex.ToString());
                return false;
            }
            return true;
        }


        /// <summary>
        /// This method gets called when an event is triggered by the monitored folder.
        /// </summary>
        /// <param name="Item">Item corresponding to the event</param>
        public static void FolderMonitorEvent(object Item)
        {
            // Only processing the MailItem
            if (Item is MailItem)
            {
                MailItem mail = (MailItem)Item;
                if (Item != null)
                {
                    Log.Out(Log.Severity.Info, "Event", "Event: New item from {0} with subject \"{1}\"!!\n", mail.SenderName, mail.Subject);
                }
            }
            else
            {
                Log.Out(Log.Severity.Info, "Event", "Event received but with unknown type " + Item.GetType().ToString());
            }
        }


        /// <summary>
        /// This method generates the recipients, either from reading the information from the passed in parameters
        /// or randomly generates it
        /// </summary>
        /// <param name="name">name of the task</param>
        /// <param name="type">type of recipients</param>
        /// <param name="recipientObject">recipient object from the Operation XML file</param>
        /// <returns>List of recipients if successful, null otherwise</returns>
        private List<string> GetRecipients(string name, RecipientTypes[] type, object[] recipientObject)
        {
            List<string> recipients = new List<string>(); ;

            // determines the recipient
            if (recipientObject == null)
            {
                Log.Out(Log.Severity.Error, name, "Recipient is not specified");
                return null;
            }

            if (type[0] == RecipientTypes.Recipients)
            {
                for (int count = 0; count < type.Length; count++)
                {
                    if (type[count] != RecipientTypes.Recipients)
                    {
                        Log.Out(Log.Severity.Error, name, "Skipping unknown recipient {0} specified", recipientObject[count].ToString());
                        continue;
                    }

                    recipients.Add((string)recipientObject[count]);
                }
            }
            // random recipients
            else
            {
                // there should only be 1 value
                if (type.Length != 1)
                {
                    Log.Out(Log.Severity.Warning, name, "More than 1 random recipients is specified, using {0}",
                        recipientObject[0].ToString());
                }

                MailSimOperationsRandomRecipients randomRecpt = (MailSimOperationsRandomRecipients)recipientObject[0];
                int randomCount = Convert.ToInt32(randomRecpt.Value);

                // query the GAL
                try
                {
                    List<string> galUsers;
                    OLAddressList gal = olMailStore.GetGlobalAddressList();

                    // uses the global distribution list if not specified
                    if (string.IsNullOrEmpty(randomRecpt.DistributionList))
                    {
                        galUsers = gal.GetUsers(null);
                    }
                    // queries the specific distribution list if specified
                    else
                    {
                        galUsers = gal.GetDLMembers(randomRecpt.DistributionList);
                    }

                    if (galUsers == null || galUsers.Count == 0)
                    {
                        throw new ArgumentException("There is no user in the GAL that matches the recipient criteria");
                    }

                    // randomly generate the number of recipients if specified
                    if (randomCount == 0)
                    {
                        // gets the number of users in GAL
                        randomCount = randomNum.Next(1, galUsers.Count + 1);
                        Log.Out(Log.Severity.Info, name, "Randomly selecting {0} recipients", randomCount);
                    }

                    // makes sure we don't pick more recipients than available
                    if (randomCount > galUsers.Count)
                    {
                        Log.Out(Log.Severity.Warning, name, "Only {0} recipients available, adjusting recipients count from {1} to {0}",
                            galUsers.Count, randomCount);
                        randomCount = galUsers.Count;
                    }

                    int recipientNumber;
                    for (int count = 0; count < randomCount; count++)
                    {
                        recipientNumber = randomNum.Next(0, galUsers.Count);
                        recipients.Add(galUsers[recipientNumber]);
                        galUsers.RemoveAt(recipientNumber);
                    }
                }
                catch (Exception ex)
                {
                    Log.Out(Log.Severity.Error, name, "Unable to get users from GAL to select random users\n" + ex.ToString());
                    return null;
                }
            }

            return recipients;
        }


        /// <summary>
        /// This method generates the attachments, either from reading the information from the passed in parameters
        /// or randomly generates it
        /// </summary>
        /// <param name="name">name of the task</param>
        /// <param name="attachmentObject">attchment object from the Operation XML file</param>
        /// <returns>List of attachments if successful, empty list otherwise</returns>
        private List<string> GetAttachments(string name, object[] attachmentObject)
        {
            List<string> attachments = new List<string>();

            // just return if no attachment is specified
            if (attachmentObject == null)
            {
                return attachments;
            }

            // determines the attachment element type
            if (attachmentObject[0].GetType() == typeof(MailSimOperationsRandomAttachments))
            {
                if (attachmentObject.Length != 1)
                {
                    Log.Out(Log.Severity.Warning, name, "More than 1 random attachements is specified, using {0}",
                        attachmentObject[0]);
                }

                MailSimOperationsRandomAttachments randomAtt = (MailSimOperationsRandomAttachments)attachmentObject[0];

                int randCount = Convert.ToInt32(randomAtt.Count);

                // makes sure the folder exists
                if (!Directory.Exists(randomAtt.Value))
                {
                    Log.Out(Log.Severity.Error, name, "Directory {0} doesn't exist, skipping attachment",
                        randomAtt.Value);
                    return attachments;
                }

                // queries all the files and randomly pick the attachment
                string[] files = Directory.GetFiles(randomAtt.Value);
                int fileNumber;

                // if Count is 0, it will attach a random number of attachments
                if (randCount == 0)
                {
                    randCount = randomNum.Next(0, files.Length);
                    Log.Out(Log.Severity.Info, name, "Randomly selecting {0} attachments", randCount);
                }

                // makes sure we don't pick more attachments than available
                if (randCount > files.Length)
                {
                    Log.Out(Log.Severity.Warning, name, "Only {0} files are available, adjusting attachment count from {1} to {0}",
                        files.Length, randCount);
                    randCount = files.Length;
                }

                for (int counter = 0; counter < randCount; counter++)
                {
                    fileNumber = randomNum.Next(0, files.Length);
                    attachments.Add(files[fileNumber]);
                }
            }
            else if (attachmentObject[0].GetType() == typeof(string))
            {
                for (int count = 0; count < attachmentObject.Length; count++)
                {
                    if (attachmentObject[count].GetType() != typeof(string))
                    {
                        Log.Out(Log.Severity.Error, name, "Skipping unknown attachment {0}. Expecting attachments path name string.",
                            attachmentObject[0].ToString());
                    }
                    else
                    {
                        attachments.Add((string)attachmentObject[count]);
                    }
                }
            }
            else
            {
                Log.Out(Log.Severity.Error, name, "Unknown attachment type {0}", attachmentObject[0].GetType());
                return null;
            }

            return attachments;
        }


        /// <summary>
        /// This method retrieves all mails or subject matching mails from the folder
        /// </summary>
        /// <param name="operationName">name of the operation</param>
        /// <param name="folder">folder to retrieve</param>
        /// <param name="subject">subject to match</param>
        /// <returns>list of mails if successful, null otherwise</returns>
        private List<MailItem> GetMails(string operationName, string folder, string subject)
        {
            List<MailItem> mails = new List<MailItem>();

            // retrieves the Outlook folder
            MailFolder mailFolder = olMailStore.GetDefaultFolder(folder);
            if (mailFolder == null)
            {
                Log.Out(Log.Severity.Error, operationName, "Unable to retrieve folder {0}",
                    folder);
                return null;
            }

            // retrieves the mail items from the folder
            MailItems folderItems = mailFolder.GetMailItems();

            if (folderItems == null || folderItems.Count == 0)
            {
                Log.Out(Log.Severity.Error, operationName, "No item in folder {0}", folder);
                return null;
            }

            // finds all mail items with matching subject if specified
            mails = FindMailWithSubject(folderItems, subject);
            if (mails.Count == 0)
            {
                Log.Out(Log.Severity.Error, operationName, "Unable to find mail with subject {0}",
                    subject);
                return null;
            }

            return mails;
        }


        /// <summary>
        /// This method returns matching subfolders of the folder.
        /// </summary>
        /// <param name="operationName">name of the operation</param>
        /// <param name="rootFolder">the root folder to retrieve subfolders</param>
        /// <param name="folderName">folder name for searching subfolders</param>
        /// <returns>list of sub folders that matches the folderName</returns>
        private List<MailFolder> GetMatchingSubFolders(string operationName, MailFolder rootFolder, string folderName)
        {
            List<MailFolder> matchingFolders = new List<MailFolder>();
            MailFolders subFolders = rootFolder.GetSubFolders();

            if (subFolders.Count == 0)
            {
                Log.Out(Log.Severity.Warning, operationName, "No subfolder in folder {0}", rootFolder.Name);
            }

            foreach (MailFolder folder in subFolders)
            {
                // we just copy all the folders to the list if folderName is not specified
                if (string.IsNullOrEmpty(folderName) || folder.Name.Contains(folderName))
                {
                    matchingFolders.Add(folder);
                }
            }

            return matchingFolders;
        }


        /// <summary>
        /// This method finds all the mails that match the given subject
        /// </summary>
        /// <param name="mails">mails to search for</param>
        /// <param name="subject">subject to match</param>
        /// <returns>list of matched mails</returns>
        private List<MailItem> FindMailWithSubject(MailItems mails, string subject)
        {
            List<MailItem> matchingMails = new List<MailItem>();

            // loops thru each mail item to find the matching subject ones
            foreach (MailItem item in mails)
            {
                if (string.IsNullOrEmpty(subject) || item.Subject.Contains(subject))
                {
                    matchingMails.Add(item);
                }
            }

            return matchingMails;
        }


        /// <summary>
        /// This method updates the registry to turn on/off Outlook prompts.
        /// This is documented in http://support.microsoft.com/kb/926512
        /// </summary>
        /// <param name="show">True to enable Outlook prompts, False to disable Outlook prompts.</param>
        public void ConfigOutlookPrompts(bool show)
        {
            const string olSecurityKey = @"HKEY_CURRENT_USER\" + OutlookPolicyRegistryRoot + @"\Security";
            const string adminSecurityMode = "AdminSecurityMode";
            const string addressBookAccess = "PromptOOMAddressBookAccess";
            const string addressInformationAccess = "PromptOOMAddressInformationAccess";
            const string saveAs = "PromptOOMSaveAs";
            const string customAction = "PromptOOMCustomAction";
            const string send = "PromptOOMSend";
            const string meetingRequestResponse = "PromptOOMMeetingTaskRequestResponse";

            try
            {
                if (show == true)
                {
                    Registry.SetValue(olSecurityKey, adminSecurityMode, (int)0);
                    Registry.SetValue(olSecurityKey, addressBookAccess, (int)1);
                    Registry.SetValue(olSecurityKey, addressInformationAccess, (int)1);
                    Registry.SetValue(olSecurityKey, customAction, (int)1);
                    Registry.SetValue(olSecurityKey, saveAs, (int)1);
                    Registry.SetValue(olSecurityKey, send, (int)1);
                    Registry.SetValue(olSecurityKey, meetingRequestResponse, (int)1);
                }
                else
                {
                    Registry.SetValue(olSecurityKey, adminSecurityMode, (int)3);
                    Registry.SetValue(olSecurityKey, addressBookAccess, (int)2);
                    Registry.SetValue(olSecurityKey, addressInformationAccess, (int)2);
                    Registry.SetValue(olSecurityKey, customAction, (int)2);
                    Registry.SetValue(olSecurityKey, saveAs, (int)2);
                    Registry.SetValue(olSecurityKey, send, (int)2);
                    Registry.SetValue(olSecurityKey, meetingRequestResponse, (int)2);
                }
            }
            catch (Exception ex)
            {
                Log.Out(Log.Severity.Error, "", "Unable to change registry, you may want to run this as Administrator\n" + ex.ToString());
            }
        }
    }
}
