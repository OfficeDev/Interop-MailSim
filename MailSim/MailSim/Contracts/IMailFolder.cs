using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSim.Contracts
{
    interface IMailFolder
    {
        /// <summary>
        /// Display Name of current folder
        /// </summary>
        string Name { get; }
        /// <summary>
        /// Folder Path of this folder
        /// </summary>
        string FolderPath { get; }
        /// <summary>
        /// Collection of MailItems in current folder
        /// </summary>
        /// <returns>IEnumerable of IMailItem</returns>
        IEnumerable<IMailItem> MailItems { get; }
        /// <summary>
        /// Collection of subfolders within this folder
        /// </summary>
        /// <returns>IEnumerable of IMailFolder</returns>
        IEnumerable<IMailFolder> SubFolders { get; }
        /// <summary>
        /// Registers event handler for ItemAdd event for new mail in the folder (i.e. new MailSim.OL.MailItem).
        /// </summary>
        /// <param name="callback">public static void FolderEvent(MailItem mail)</param>
        void RegisterItemAddEventHandler(Action<IMailItem> callback);
        /// <summary>
        /// Unregisters event handler previously registered with RegisterItemAddEventHandler
        /// </summary>
        void UnRegisterItemAddEventHandler();
        /// <summary>
        /// Number of mail items in the current folder
        /// </summary>
        int MailItemsCount { get; }
        /// <summary>
        /// Number of subfolders in the current folder
        /// </summary>
        int SubFoldersCount { get; }
        /// <summary>
        /// Adds folder as a subfolder of current folder
        /// </summary>
        /// <param name="name">Nam of new folder</param>
        /// <returns></returns>
        IMailFolder AddSubFolder(string name);
        /// <summary>
        /// Deletes current folder
        /// </summary>
        void Delete();
#if false
        /// <summary>
        /// Default Message Class of current folder
        /// </summary>
        public string DefaultMessageClass { get; }
#endif
    }
}
