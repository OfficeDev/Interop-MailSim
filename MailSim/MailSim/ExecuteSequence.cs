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
using System.Linq;

using MailSim.Contracts;

namespace MailSim
{
    class ExecuteSequence
    {
        private MailSimSequence sequence;
        private MailSimOperations operations;
        private XmlDocument operationXML;
        private IMailStore olMailStore;
        private Random randomNum;

        private const string OfficeVersion = "15.0";
        private const string OfficePolicyRegistryRoot = @"Software\Policies\Microsoft\Office\" + OfficeVersion;
        private const string OutlookPolicyRegistryRoot = OfficePolicyRegistryRoot + @"\Outlook";

        private const string Recipients = "Recipients";
        private const string RandomRecipients = "RandomRecipients";
        private const string DefaultSubject = "Default Subject";
        private const string DefaultBody = "Default Body";
        private const int MaxNumberOfRandomFolder = 100;
        private const string StopFileName = "stop.txt";

        private List<IMailFolder> FolderEventList = new List<IMailFolder>();
        private IDictionary<Type, Func<object, bool>> typeFuncs = new Dictionary<Type, Func<object, bool>>();

        private string DefaultInboxMonitor = "DefaultInboxMonitor";
        public static string eventString = "Event";


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="seq">Sequence file content </param>
        public ExecuteSequence(MailSimSequence seq)
        {
            typeFuncs[typeof(MailSimOperationsMailSend)] = (oper) => MailSend((MailSimOperationsMailSend)oper);
            typeFuncs[typeof(MailSimOperationsMailDelete)] = (oper) => MailDelete((MailSimOperationsMailDelete)oper);
            typeFuncs[typeof(MailSimOperationsMailReply)] = (oper) => MailReply((MailSimOperationsMailReply)oper);
            typeFuncs[typeof(MailSimOperationsMailForward)] = (oper) => MailForward((MailSimOperationsMailForward)oper);
            typeFuncs[typeof(MailSimOperationsMailMove)] = (oper) => MailMove((MailSimOperationsMailMove)oper);
            typeFuncs[typeof(MailSimOperationsFolderCreate)] = (oper) => FolderCreate((MailSimOperationsFolderCreate)oper);
            typeFuncs[typeof(MailSimOperationsFolderDelete)] = (oper) => FolderDelete((MailSimOperationsFolderDelete)oper);
            typeFuncs[typeof(MailSimOperationsEventMonitor)] = (oper) => EventMonitor((MailSimOperationsEventMonitor)oper);

            if (seq != null)
            {
                try
                {
                    sequence = seq;

                    // Disables the Outlook security prompt if specified
                    if (sequence.DisableOutlookPrompt == true)
                    {
                        ConfigOutlookPrompts(false);
                    }

                    // Openes connection to Outlook with default profile, starts Outlook if it is not running
                    // Note: Currently only the default MailStore is supported.
                    olMailStore = ProviderFactory.CreateMailStore(null);

                    // Initializes a random number
                    randomNum = new Random();
                }
                catch (Exception)
                {
                    Log.Out(Log.Severity.Error, "Run", "Error encountered during initialization");
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

            // Restore the Outlook prompt if needed 
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
            // Unregisters all registered folder events 
            foreach (IMailFolder folder in FolderEventList)
            {
                RegisterFolderEvent(DefaultInboxMonitor, folder, false);
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

            // Registers to monitor the Inbox
            MailSimOperationsEventMonitor inboxEvent = new MailSimOperationsEventMonitor();
            inboxEvent.Folder = "olFolderInbox";
            inboxEvent.OperationName = "DefaultInboxMonitor";
            EventMonitor(inboxEvent);

            // Run each operation group
            foreach (MailSimSequenceOperationGroup group in sequence.OperationGroup)
            {
                int iterations = GetIterationCount(group.Iterations);

                // Run the operations file
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

                    Log.Out(Log.Severity.Info, group.Name, "Completed group task run {0}", count);

                    SleepOrStop(group.Name, group.Sleep);
                }

                CleanupAfterIteration();
            }
        }


        /// <summary>
        /// This method runs each task of the OperationGroup, handling the iteration and sleep elements.
        /// </summary>
        /// <param name="task">task to run</param>
        /// <returns>Returns true if successful, otherwise returns false </returns>
        public void ProcessTask(MailSimSequenceOperationGroupTask task)
        {
            int iterations = GetIterationCount(task.Iterations);

            for (int count = 1; count <= iterations; count++)
            {
                Log.Out(Log.Severity.Info, task.Name, "Running task {0}", count);

                if (ExecuteTask(task.Name))
                {
                    Log.Out(Log.Severity.Info, task.Name, "Completed task run {0}", count);
                }
                else
                {
                    Log.Out(Log.Severity.Error, task.Name, "Failed to run task");
                }

                SleepOrStop(task.Name, task.Sleep);
            }
        }

        /// <summary>
        /// This method determines and calls the appropriate method to run the task
        /// </summary>
        /// <param name="taskName">name of the task</param>
        /// <returns>Returns true if successful, otherwise returns false </returns>
        public bool ExecuteTask(string taskName)
        {
            object operation = null;

            try
            {
                operation = operations.Items.SingleOrDefault(x => GetOperationName(x) == taskName);
            }
            catch (InvalidOperationException)
            {
                // we expect only 1 operation node matching the name
                Log.Out(Log.Severity.Error, taskName, "More than one task with this name");
                return false;
            }

            if (operation == null)
            {
                Log.Out(Log.Severity.Error, taskName, "Unable to find matching task; skipping task");
                return false;
            }

            Func<object, bool> func;

            if (typeFuncs.TryGetValue(operation.GetType(), out func) == false)
            {
                Log.Out(Log.Severity.Error, taskName, "Undefined task; skipping");
                return false;
            }

            return func(operation);
        }

        /// <summary>
        /// This method sends mail according to the parameter
        /// </summary>
        /// <param name="operation">parameters for MailSend</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailSend(MailSimOperationsMailSend operation)
        {
            int iterations = GetIterationCount(operation.Count);
 
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
                    Log.Out(Log.Severity.Error, operation.OperationName, "Recipient is not specified, skipping the operation");
                    return false;
                }

                List<string> attachments = GetAttachments(operation.OperationName, operation.Attachments);

                try
                {
                    // generates a new email
                    IMailItem mail = olMailStore.NewMailItem();

                    mail.Subject = mail.Body = System.DateTime.Now.ToString() + " - ";
                    mail.Subject += (string.IsNullOrEmpty(operation.Subject)) ? DefaultSubject : operation.Subject;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Subject: {0}", mail.Subject);
                    mail.Body += (string.IsNullOrEmpty(operation.Body)) ? DefaultBody : operation.Body;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Body: {0}", mail.Body);

                    // Adds all recipients
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

                SleepOrStop(operation.OperationName, operation.Sleep);
            }

            return true;
        }


        /// <summary>
        /// This method deletes mail according to the parameter 
        /// </summary>
        /// <param name="operation">parameters for MailDelete</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailDelete(MailSimOperationsMailDelete operation)
        {
            int iterations = GetIterationCount(operation.Count);
            bool random = false;

            try
            {
                // Retrieves mails from Outlook
                var mails = GetMails(operation.OperationName, operation.Folder, operation.Subject).ToList();
                if (mails.Any() == false)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Skipping MailDelete");
                    return false;
                }

                // Randomly generate the number of emails to delete 
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
                        "Only {1} email(s) are in the folder, so the number of emails to delete is adjusted from {0} to {1}",
                        iterations, mails.Count);
                    iterations = mails.Count;
                }

                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, operation.OperationName, "Starting iteration {0}", count);

                    // just delete the email in order if random is not selected,
                    // otherwise randomly pick the mail to delete
                    int indexToDelete = random ? randomNum.Next(0, mails.Count) : mails.Count - 1;

                    Log.Out(Log.Severity.Info, operation.OperationName, "Deleting email with subject: {0}", mails[indexToDelete].Subject);
                    mails[indexToDelete].Delete();
                    mails.RemoveAt(indexToDelete);

                    SleepOrStop(operation.OperationName, operation.Sleep);

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
            int iterations = GetIterationCount(operation.Count);
            bool random = false;

            try
            {
                // retrieves mails from Outlook
                var mails = GetMails(operation.OperationName, operation.Folder, operation.MailSubjectToReply).ToList();
                if (mails.Any() == false)
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
                        "Only {1} email(s) are in the folder, so the number of emails to reply is adjusted from {0} to {1}",
                        iterations, mails.Count);
                    iterations = mails.Count;
                }

                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, operation.OperationName, "Starting iteration {0}", count);

                    List<string> attachments = GetAttachments(operation.OperationName, operation.Attachments);

                    // just reply the email in order if random is not selected,
                    // otherwise randomly pick the mail to reply
                    int indexToReply = random ? randomNum.Next(0, mails.Count) : count - 1;
                    IMailItem mailToReply = mails[indexToReply].Reply(operation.ReplyAll);

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

                    SleepOrStop(operation.OperationName, operation.Sleep);

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
        /// <param name="operation">argument for MailForward</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailForward(MailSimOperationsMailForward operation)
        {
            int iterations = GetIterationCount(operation.Count);
            bool random = false;
 
            try
            {
                // retrieves mails from Outlook
                var mails = GetMails(operation.OperationName, operation.Folder, operation.MailSubjectToForward).ToList();
                if (mails.Any() == false)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Skipping MailForward");
                    return false;
                }

                // randomly generates the number of emails to forward 
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
                        "Only {1} email(s) are in the folder, so the the number of emails to forward is adjusted from {0} to {1}",
                        iterations, mails.Count);
                    iterations = mails.Count;
                }

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
                    int indexToForward = random ? randomNum.Next(0, mails.Count) : count - 1;
                    IMailItem mailToForward = mails[indexToForward].Forward();

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

                    SleepOrStop(operation.OperationName, operation.Sleep);

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
        /// <param name="operation">argument for MaiMove</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailMove(MailSimOperationsMailMove operation)
        {
            int iterations = GetIterationCount(operation.Count);
            bool random = false;

            try
            {
                // retrieves mails from Outlook
                var mails = GetMails(operation.OperationName, operation.SourceFolder, operation.Subject).ToList();
                if (mails.Any() == false)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Skipping MailMove");
                    return false;
                }

                // randomly generates the number of emails to forward 
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
                        "Only {1} email(s) are in the folder, so the number of emails to move is adjusted from {0} to {1}",
                        iterations, mails.Count);
                    iterations = mails.Count;
                }

                // gets the Outlook destination folder
                IMailFolder destinationFolder = olMailStore.GetDefaultFolder(operation.DestinationFolder);
                if (destinationFolder == null)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Unable to retrieve folder {0}",
                        operation.DestinationFolder);
                    return false;
                }

                for (int count = 1; count <= iterations; count++)
                {
                    Log.Out(Log.Severity.Info, operation.OperationName, "Starting iteration {0}", count);

                    // just move the email in order if random is not selected,
                    // otherwise randomly pick the mail to move
                    int indexToCopy = random ? randomNum.Next(0, mails.Count) : mails.Count - 1;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Moving to {0}: {1}",
                        operation.DestinationFolder, mails[indexToCopy].Subject);

                    mails[indexToCopy].Move(destinationFolder);
                    mails.RemoveAt(indexToCopy);

                    SleepOrStop(operation.OperationName, operation.Sleep);

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
            int iterations = GetIterationCount(operation.Count);

            try
            {
                // gets the Outlook folder
                IMailFolder folder = olMailStore.GetDefaultFolder(operation.FolderPath);
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
                    string newFolderName = System.DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + " - " + operation.FolderName;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Creating folder: {0}", newFolderName);
                    folder.AddSubFolder(newFolderName);

                    SleepOrStop(operation.OperationName, operation.Sleep);

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
            int iterations = GetIterationCount(operation.Count);
            bool random = false;

            try
            {
                // gets the Outlook folder
                IMailFolder folder = olMailStore.GetDefaultFolder(operation.FolderPath);
                if (folder == null)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Unable to retrieve folder {0}",
                        operation.FolderPath);
                    return false;
                }

                var subFolders = GetMatchingSubFolders(operation.OperationName, folder, operation.FolderName).ToList();
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

                    SleepOrStop(operation.OperationName, operation.Sleep);
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
                IMailFolder folder = olMailStore.GetDefaultFolder(operation.Folder);
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

                SleepOrStop(operation.OperationName, operation.Sleep);
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
        private bool RegisterFolderEvent(string operation, IMailFolder folder, bool register)
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
            if (Item == null)
            {
                Log.Out(Log.Severity.Info, eventString, "Unknown event received");
                return;
            }

            // Only processing the MailItem
            if (Item is IMailItem)
            {
                IMailItem mail = (IMailItem)Item;
                Log.Out(Log.Severity.Info, eventString, "New item from {0} with subject \"{1}\"!!", mail.SenderName, mail.Subject);
            }
            else
            {
                Log.Out(Log.Severity.Info, eventString, "Event received but with unknown type " + Item.GetType().ToString());
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

                    var gal = olMailStore.GetGlobalAddressList();

                    // uses the global distribution list if not specified
                    if (string.IsNullOrEmpty(randomRecpt.DistributionList))
                    {
                        galUsers = gal.GetUsers(null).ToList();
                    }
                    // queries the specific distribution list if specified
                    else
                    {
                        galUsers = gal.GetDLMembers(randomRecpt.DistributionList).ToList();
                    }

                    if (galUsers.Any() == false)
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

                    for (int count = 0; count < randomCount; count++)
                    {
                        int recipientNumber = randomNum.Next(0, galUsers.Count);
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
            }

            return attachments;
        }


        /// <summary>
        /// This method retrieves all mails or subject matching mails from the folder
        /// </summary>
        /// <param name="operationName">name of the operation</param>
        /// <param name="folder">folder to retrieve</param>
        /// <param name="subject">case sensitive subject to match</param>
        /// <returns>list of mails if successful, null otherwise</returns>
        private IEnumerable<IMailItem> GetMails(string operationName, string folder, string subject)
        {
            // retrieves the Outlook folder
            IMailFolder mailFolder = olMailStore.GetDefaultFolder(folder);
            if (mailFolder == null)
            {
                Log.Out(Log.Severity.Error, operationName, "Unable to retrieve folder {0}",
                    folder);
                return Enumerable.Empty<IMailItem>();
            }

            // retrieves the mail items from the folder
            var mails = mailFolder.MailItems;

            if (mails.Any() == false)
            {
                Log.Out(Log.Severity.Error, operationName, "No item in folder {0}", folder);
                return mails;
            }

            // finds all mail items with matching subject if specified
            subject = subject ?? string.Empty;
            mails = mails.Where(x => x.Subject.Contains(subject));

            if (mails.Any() == false)
            {
                Log.Out(Log.Severity.Error, operationName, "Unable to find mail with subject {0}",
                    subject);
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
        private IEnumerable<IMailFolder> GetMatchingSubFolders(string operationName, IMailFolder rootFolder, string folderName)
        {
            var subFolders = rootFolder.SubFolders;

            if (subFolders.Any() == false)
            {
                Log.Out(Log.Severity.Warning, operationName, "No subfolders in folder {0}", rootFolder.Name);
            }

            folderName = folderName ?? string.Empty;

            return subFolders.Where(x => x.Name.Contains(folderName));
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

        private int GetIterationCount(string countString)
        {
            if (string.IsNullOrEmpty(countString))
            {
                return 1;
            }

            return Convert.ToInt32(countString);
        }

        private string GetOperationName(object operation)
        {
            dynamic op = operation;

            return op.OperationName;
        }

        private void SleepOrStop(string name, string sleepSeconds)
        {
            if (File.Exists(StopFileName))
            {
                Log.Out(Log.Severity.Info, string.Format("StopApplication at {0}", name), "Stopping simulation run...");
                Environment.Exit(0);
            }

            if (!string.IsNullOrEmpty(sleepSeconds))
            {
                int sleep = Convert.ToInt32(sleepSeconds);
                Log.Out(Log.Severity.Info, name, "Sleeping for {0} seconds", sleep);
                Thread.Sleep(sleep * 1000);
            }
        }
    }
}
