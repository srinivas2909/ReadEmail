using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace EmailAttachments
{
    class Program
    {
        static void Main(string[] args)
        {
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;

            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;
                string Filter = "[ReceivedTime] >= Today";

                Items mis = inboxFolder.Items.Restrict(Filter);
                int cnt = mis.Count; ;

                Console.WriteLine(mis.Count);

                foreach (MailItem item in mis)
                {
                

                    if (item.Attachments.Count > 0)
                    {
                        for (int i = 1; i <= item.Attachments.Count; i++)
                        {
                            item.Attachments[i].SaveAsFile
                                (@"C:\TestFileSave\" +
                                item.Attachments[i].FileName);
                        }
                    }
                    //Console.WriteLine(stringBuilder);
                    Marshal.ReleaseComObject(item);
                }
            }
            //Error handler.
            catch (System.Exception e)
            {
                Console.WriteLine("{0} Exception caught: ", e);
            }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }

            Console.WriteLine("OK");
            Console.ReadKey();
        }
        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
