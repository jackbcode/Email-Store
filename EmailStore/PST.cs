using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailStore
{
   public static class PST
    {
        public static IEnumerable<MailItem> readPst(string pstFilePath, string pstName)
        {

            Console.WriteLine(" readpst started");
            List<MailItem> mailItems = new List<MailItem>();

            Console.WriteLine("started app");
            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            Console.WriteLine(" get namespace");
            NameSpace outlookNs = app.GetNamespace("MAPI");
            Console.WriteLine(" addstore");
            //Add PST file(Outlook Data File) to Default Profile
            outlookNs.AddStore(pstFilePath);

            string storeInfo = null;

            //foreach (Store store in outlookNs.Stores)
            //{
            //    storeInfo = store.DisplayName;
            //    storeInfo = store.FilePath;
            //    storeInfo = store.StoreID;
            //}

            Console.WriteLine("getting root folder");
            MAPIFolder rootFolder = outlookNs.Stores[pstName].GetRootFolder();
            //MAPIFolder rootFolder = outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            Console.WriteLine("getting subfolders");
            // Traverse through all folders in the PST file
            Folders subFolders = rootFolder.Folders;

            foreach (Folder folder in subFolders)
            {
                ExtractItems(mailItems, folder);
           
            }
            // Remove PST file from Default Profile
            outlookNs.RemoveStore(rootFolder);
            return mailItems;
        }

        public static void ExtractItems(List<MailItem> mailItems, Folder folder)
        {
            Items items = folder.Items;

            int itemcount = items.Count;

            foreach (object item in items)
            {
                if (item is MailItem)
                {
                   
                    MailItem mailItem = item as MailItem;
                    Console.WriteLine("Adding " + mailItem.Subject);
                    mailItems.Add(mailItem);
                }
            }

            foreach (Folder subfolder in folder.Folders)
            {
                ExtractItems(mailItems, subfolder);
            }
        }
    }
}
