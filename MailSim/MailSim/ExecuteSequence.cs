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
using System.Linq;

using MailSim.Common;
using MailSim.Common.Contracts;

namespace MailSim
{
    class ExecuteSequence
    {
        private readonly MailSimSequence sequence;
        private MailSimOperations operations;
        private XmlDocument operationXML;
        private readonly IMailStore olMailStore;
        private readonly Random randomNum = new Random();

        private const string Recipients = "Recipients";
        private const string RandomRecipients = "RandomRecipients";
        private const string DefaultSubject = "Default Subject";
        private const string DefaultBody = "Default Body";
        private const int MaxNumberOfRandomFolder = 100;
        private const string StopFileName = "stop.txt";

        private readonly List<IMailFolder> FolderEventList = new List<IMailFolder>();
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

            sequence = seq;

            if (sequence != null)
            {
                try
                {
                    olMailStore = ProviderFactory.CreateMailStore(null, sequence);
                }
                catch (Exception)
                {
                    Log.Out(Log.Severity.Error, "Run", "Error encountered during initialization");
                    throw;
                }
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
                operation = operations.Items
                    .SingleOrDefault(x => string.Equals(GetOperationName(x), taskName, StringComparison.OrdinalIgnoreCase));
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
                try
                {
                    // generates a new email
                    IMailItem mail = olMailStore.NewMailItem();

                    mail.Subject = DateTime.Now.ToString() + " - ";
                    mail.Subject += (string.IsNullOrEmpty(operation.Subject)) ? DefaultSubject : operation.Subject;
                    Log.Out(Log.Severity.Info, operation.OperationName, "Subject: {0}", mail.Subject);

                    mail.Body = BuildBody(operation.Body);
                    Log.Out(Log.Severity.Info, operation.OperationName, "Body: {0}", mail.Body);

                    if (!AddRecipients(mail, operation))
                    {
                        return false;
                    }

                    AddAttachments(mail, operation);

                    mail.Send();
                }
                catch (Exception ex)
                {
                    Log.Out(Log.Severity.Error, operation.OperationName, "Exception encountered\n{0}", ex);
                    return false;
                }

                SleepOrStop(operation.OperationName, operation.Sleep);
            }

            return true;
        }


        private ParsedOperation ParseOperation(dynamic op, string folder, string subject)
        {
            // Retrieves mails from Outlook
            IEnumerable<IMailItem> mails = GetMails(op, folder, subject);

            return new ParsedOperation(op, mails.ToList());
        }

        /// <summary>
        /// This method deletes mail according to the parameter 
        /// </summary>
        /// <param name="operation">parameters for MailDelete</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailDelete(MailSimOperationsMailDelete operation)
        {
            var parsedOp = ParseOperation(operation, operation.Folder, operation.Subject);

            return parsedOp.Iterate((indexToDelete, mails) =>
            {
                var item = mails[indexToDelete];
                Log.Out(Log.Severity.Info, operation.OperationName, "Deleting email with subject: \"{0}\"", item.Subject);
                
                item.Delete();
                mails.RemoveAt(indexToDelete);
                return true;
            });
        }

        /// <summary>
        /// This method replies email according to the parameters
        /// </summary>
        /// <param name="operation">parameters for MailReply</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailReply(MailSimOperationsMailReply operation)
        {
            var parsedOp = ParseOperation(operation, operation.Folder, operation.MailSubjectToReply);

            return parsedOp.Iterate((indexToReply, mails) =>
            {
                IMailItem mailToReply = mails[indexToReply].Reply(operation.ReplyAll);

                Log.Out(Log.Severity.Info, operation.OperationName, "Subject: {0}", mailToReply.Subject);

                mailToReply.Body = BuildBody(operation.ReplyBody) + mailToReply.Body;
                Log.Out(Log.Severity.Info, operation.OperationName, "Body: {0}", mailToReply.Body);

                // process the attachment
                AddAttachments(mailToReply, operation);

                mailToReply.Send();
                return true;
            });
        }

        /// <summary>
        /// This method forwards emails according to the parameters
        /// </summary>
        /// <param name="operation">argument for MailForward</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailForward(MailSimOperationsMailForward operation)
        {
            var parsedOp = ParseOperation(operation, operation.Folder, operation.MailSubjectToForward);

            return parsedOp.Iterate((indexToForward, mails) =>
            {
                IMailItem mailToForward = mails[indexToForward].Forward();

                Log.Out(Log.Severity.Info, operation.OperationName, "Subject: {0}", mailToForward.Subject);

                mailToForward.Body = BuildBody(operation.ForwardBody) + mailToForward.Body;

                Log.Out(Log.Severity.Info, operation.OperationName, "Body: {0}", mailToForward.Body);

                if (!AddRecipients(mailToForward, operation/*, operation.Items[0] is MailSimOperationsMailForwardRandomAttachments*/))
                {
                    return false;
                }

                AddAttachments(mailToForward, operation);

                mailToForward.Send();
                return true;
            });
        }

        /// <summary>
        /// This method moves emails according to the parameters
        /// </summary>
        /// <param name="operation">argument for MaiMove</param>
        /// <returns>true if processed successfully, false otherwise</returns>
        private bool MailMove(MailSimOperationsMailMove operation)
        {
            // gets the Outlook destination folder
            IMailFolder destinationFolder = olMailStore.GetDefaultFolder(operation.DestinationFolder);
            if (destinationFolder == null)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Unable to retrieve folder {0}",
                    operation.DestinationFolder);
                return false;
            }

            var parsedOp = ParseOperation(operation, operation.SourceFolder, operation.Subject);

            return parsedOp.Iterate((indexToMove, mails) =>
            {
                var item = mails[indexToMove];
                Log.Out(Log.Severity.Info, operation.OperationName, "Moving to {0}: {1}",
                    operation.DestinationFolder,item.Subject);

                item.Move(destinationFolder);
                mails.RemoveAt(indexToMove);
                return true;
            });
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
                    string newFolderName = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss.fff") + " - " + operation.FolderName;
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
                Log.Out(Log.Severity.Error, operation, "RegisterFolderEvent: Exception encountered\n{0}", ex);
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
        private List<string> GetRecipients(dynamic operation)
        {
            string name = operation.OperationName;

            string[] specificRecipients = operation.Recipient;
            RandomRecipients randomRecipients = operation.RandomRecipients;

            // determines the recipient
            if (specificRecipients == null && randomRecipients == null)
            {
                Log.Out(Log.Severity.Error, name, "Recipients are not specified");
                return null;
            }

            List<string> recipientNames = new List<string>();

            if (specificRecipients != null)
            {
                recipientNames.AddRange(specificRecipients);
            }

            if (randomRecipients != null)
            {
                int randomCount = Convert.ToInt32(randomRecipients.Value);

                // query the GAL
                try
                {
                    List<string> galUsers;

                    var gal = olMailStore.GetGlobalAddressList();
                    int userCountForRandom = int.Parse(operation.UserCountForRandomization);

                    // uses the global distribution list if not specified
                    if (string.IsNullOrEmpty(randomRecipients.DistributionList))
                    {
                        galUsers = gal.GetUsers(null, userCountForRandom).ToList();
                    }
                    // queries the specific distribution list if specified
                    else
                    {
                        galUsers = gal.GetDLMembers(randomRecipients.DistributionList, userCountForRandom).ToList();
                    }

                    if (galUsers.Any() == false)
                    {
                        throw new ArgumentException("There are no users in the GAL that matches the recipient criteria");
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
                        recipientNames.Add(galUsers[recipientNumber]);
                        galUsers.RemoveAt(recipientNumber);
                    }
                }
                catch (Exception ex)
                {
                    Log.Out(Log.Severity.Error, name, "Unable to get users from GAL to select random users\n{0}", ex);
                    return null;
                }
            }

            return recipientNames;
        }


        /// <summary>
        /// This method generates the attachments, either from reading the information from the passed in parameters
        /// or randomly generates it
        /// </summary>
        /// <param name="name">name of the task</param>
        /// <param name="attachmentObject">attchment object from the Operation XML file</param>
        /// <returns>List of attachments if successful, empty list otherwise</returns>
        private List<string> GetAttachments(dynamic operation)
        {
            List<string> attachments = new List<string>();
            string name = operation.OperationName;

            string[] specificAttachments = operation.Attachment;
            RandomAttachments randomAttachments = operation.RandomAttachments;

            if (specificAttachments != null)
            {
                attachments.AddRange(specificAttachments);
            }

            if (randomAttachments != null)
            {
                int randCount = Convert.ToInt32(randomAttachments.Count);
                string dir = randomAttachments.Value.Trim();

                // makes sure the folder exists
                if (!Directory.Exists(dir))
                {
                    Log.Out(Log.Severity.Error, name, "Directory {0} doesn't exist, skipping attachment",
                        randomAttachments.Value);
                    return attachments;
                }

                // queries all the files and randomly pick the attachment
                var files = Directory.GetFiles(dir).ToList();
                int fileNumber;

                // if Count is 0, it will attach a random number of attachments
                if (randCount == 0)
                {
                    randCount = randomNum.Next(0, files.Count);
                    Log.Out(Log.Severity.Info, name, "Randomly selecting {0} attachments", randCount);
                }

                // makes sure we don't pick more attachments than available
                if (randCount > files.Count)
                {
                    Log.Out(Log.Severity.Warning, name, "Only {0} files are available, adjusting attachment count from {1} to {0}",
                        files.Count, randCount);
                    randCount = files.Count;
                }

                for (int counter = 0; counter < randCount; counter++)
                {
                    fileNumber = randomNum.Next(0, files.Count);
                    attachments.Add(files[fileNumber]);
                    files.RemoveAt(fileNumber);
                }
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
        private IEnumerable<IMailItem> GetMails(dynamic op, string folder, string subject)
        {
            // retrieves the Outlook folder
            IMailFolder mailFolder = olMailStore.GetDefaultFolder(folder);
            if (mailFolder == null)
            {
                Log.Out(Log.Severity.Error, op.OperationName, "Unable to retrieve folder {0}",
                    folder);
                return Enumerable.Empty<IMailItem>();
            }

            // retrieves the mail items from the folder
            int mailCountForRandom = int.Parse(op.MailCountForRandomization);

            var mails = mailFolder.GetMailItems(subject, mailCountForRandom);

            if (mails.Any() == false)
            {
                Log.Out(Log.Severity.Error, op.OperationName, "No items with subject \"{0}\" in folder {1}", subject, folder);
                return mails;
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

        private bool AddRecipients(IMailItem mail, dynamic operation)
        {
            var recipients = GetRecipients(operation);

            if (recipients == null)
            {
                Log.Out(Log.Severity.Error, operation.OperationName, "Recipients are not specified, skipping operation");
                return false;
            }

            // Add all recipients
            foreach (string recpt in recipients)
            {
                Log.Out(Log.Severity.Info, operation.OperationName, "Recipient: {0}", recpt);
                mail.AddRecipient(recpt);
            }

            return true;
        }

        private void AddAttachments(IMailItem mail, dynamic operation)
        {
            var attachments = GetAttachments(operation);

            foreach (string attmt in attachments)
            {
                Log.Out(Log.Severity.Info, operation.OperationName, "Attachment: {0}", attmt);
                mail.AddAttachment(attmt);
            }
        }

        private string BuildBody(string templateBody)
        {
            string body = DateTime.Now.ToString() + " - " +
                                ((string.IsNullOrEmpty(templateBody)) ? DefaultBody : templateBody);
            return body;
        }

        private static int GetIterationCount(string countString)
        {
            return string.IsNullOrEmpty(countString) ? 1 : Convert.ToInt32(countString);
        }

        private string GetOperationName(dynamic operation)
        {
            return operation.OperationName;
        }

        private static void SleepOrStop(string name, string sleepSeconds)
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

        private class ParsedOperation
        {
            private const string TypePrefix = "MailSim.MailSimOperations";
            private readonly Random _random = new Random();

            internal ParsedOperation(dynamic operation, IList<IMailItem> mails)
            {
                Mails = mails;
                Op = operation;

                InitIterations();
            }

            private dynamic Op { get; set; }
            private IList<IMailItem> Mails { get; set; }
            private int Iterations { get; set; }
            private bool IsRandom { get; set; }

            private int GetNextIndex()
            {
                return IsRandom ? _random.Next(0, Mails.Count) : Mails.Count - 1;
            }

            private void InitIterations()
            {
                var name = Op.GetType().ToString();
                name = name.Substring(TypePrefix.Length);
                int mailCount = Mails.Count;

                if (mailCount == 0)
                {
                    Log.Out(Log.Severity.Error, Op.OperationName, "Skipping " + Op.OperationName);
                    Iterations = 0;
                }
                else
                {
                    Iterations = GetIterationCount(Op.Count);
                    IsRandom = Iterations == 0;

                    // Randomly generate the number of emails
                    if (IsRandom)
                    {
                        Iterations = _random.Next(1, mailCount + 1);
                        Log.Out(Log.Severity.Info, Op.OperationName, "Randomly applying {0} to {1} emails", Op.OperationName, Iterations);
                    }
                    // we need to make sure we are not deleting more than what we have in the mailbox
                    else if (Iterations > mailCount)
                    {
                        Log.Out(Log.Severity.Warning, Op.OperationName,
                            "Only {1} email(s) are found, so the number of emails to {2} is adjusted from {0} to {1}",
                            Iterations, mailCount, name);
                        Iterations = mailCount;
                    }
                }
            }

            internal bool Iterate(Func<int, IList<IMailItem>, bool> func)
            {
                for (int count = 1; count <= Iterations; count++)
                {
                    Log.Out(Log.Severity.Info, Op.OperationName, "Starting iteration {0}", count);

                    try
                    {
                        if (!func(GetNextIndex(), Mails))
                        {
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Out(Log.Severity.Error, Op.OperationName, "Exception encountered\n{0}", ex);
                        return false;
                    }

                    SleepOrStop(Op.OperationName, Op.Sleep);
                    Log.Out(Log.Severity.Info, Op.OperationName, "Finished iteration {0}", count);
                }

                return true;
            }
        }
    }
}
