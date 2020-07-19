using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace Carbuncle
{
    class Program
    {
        static void Main(string[] args)
        {

            /**
             * Commands:
             *  Read(string Subject) or Read(int Num)
             *  Search(string Keyword)
             *  Enumerate()
             *  Send()
             *  */
            var parsed = ArgumentParser.Parse(args);
            if (!parsed.ParsedOk)
            {
                return;
            }
            string action = args.Length != 0 ? args[0] : "";

            switch (action.ToLower())
            {
                case "read":
                    //Read from Inbox Number or Subject
                    if (parsed.Arguments.ContainsKey("number"))
                    {
                        ReadEmail(int.Parse(parsed.Arguments["number"]));
                    }
                    else if (parsed.Arguments.ContainsKey("number"))
                    {
                        ReadEmail(parsed.Arguments["subject"]);
                    }
                    else
                    {
                        PrintHelp();
                    }
                    break;
                case "search":
                    //Search Body and Subject for Keyword
                    if (parsed.Arguments.ContainsKey("keyword"))
                    {
                        SearchAll(parsed.Arguments["keyword"]);
                    }
                    else
                    {
                        SearchAll("");
                    }
                    break;
                case "enum":
                    //List all Subjects for MailItems
                    SearchAll("");
                    break;
                case "monitor":
                    MonitorEmail();
                    while (true)
                    {

                    }
                    break;
                case "send":
                    SendEmail("scottie.austin1991@gmail.com", "Test E-mail from interop", "Interop Test");
                    break;
            }
        }
        static void PrintHelp()
        {
            Console.WriteLine("halp");
        }
        static void ReadEmail(string Subject)
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            try
            {
                foreach (var item in mailItems)
                {
                    switch (TypeInformation.GetTypeName(item))
                    {
                        case "MailItem":
                            {
                                MailItem itemCur = (MailItem)item;
                                if (itemCur.Subject.Contains(Subject))
                                {
                                    Console.WriteLine(itemCur.Body);
                                }
                                break;
                            }
                        case "MeetingItem":
                            {
                                MeetingItem itemCur = (MeetingItem)item;
                                if (itemCur.Subject.Contains(Subject))
                                {
                                    Console.WriteLine(itemCur.Body);
                                }
                                break;
                            }
                    }
                }
            }
            catch
            {
                //Console.WriteLine("Error");
            }
        }
        static void ReadEmail(int number)
        {

            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);

            MailItem item = (MailItem)mailItems[number];
            Console.WriteLine(item.Subject);
        }
        static void SearchAll(string keyword)
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            try
            {
                foreach (var item in mailItems)
                {
                    switch (TypeInformation.GetTypeName(item))
                    {
                        case "MailItem":
                            {
                                MailItem itemCur = (MailItem)item;
                                if (keyword == "" || itemCur.Body.ToLower().Contains(keyword.ToLower()) || itemCur.Body.ToLower().Contains(keyword.ToLower()))
                                {
                                    Console.WriteLine(itemCur.Subject);
                                    Console.WriteLine(itemCur.Body);
                                    Console.WriteLine();
                                }
                                break;
                            }
                        case "MeetingItem":
                            {
                                MailItem itemCur = (MailItem)item;
                                if (keyword == "" || itemCur.Body.ToLower().Contains(keyword.ToLower()) || itemCur.Body.ToLower().Contains(keyword.ToLower()))
                                {
                                    Console.WriteLine(itemCur.Subject);
                                    Console.WriteLine(itemCur.Body);
                                    Console.WriteLine();
                                }
                                break;
                            }
                    }
                }
            }
            catch
            {
                //Console.WriteLine("Error");
            }
        }
        static void SearchMeetings()
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            List<MailItem> ReceivedEmail = new List<MailItem>();
            try
            {
                foreach (var item in mailItems)
                {
                    if (TypeInformation.GetTypeName(item) == "MeetingItem")
                    {
                        MailItem itemCur = (MailItem)item;

                        Console.WriteLine("[+] " + itemCur.Subject + " - " + itemCur.ReminderTime);
                    }
                    switch (TypeInformation.GetTypeName(item))
                    {
                        case "MailItem":
                            {
                                MailItem itemCur = (MailItem)item;
                                Console.WriteLine("[+] " + itemCur.Subject);
                                break;
                            }
                        case "MeetingItem":
                            {
                                MeetingItem itemCur = (MeetingItem)item;
                                Console.WriteLine("[+] " + itemCur.Subject);
                                break;
                            }
                    }
                }
            }
            catch
            {
            }
        }
        static void MonitorEmail()
        {
            Console.WriteLine("Beginning Monitor");
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            mailItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(NewEmailEvent);
        }
        static void NewEmailEvent(object item)
        {
            Console.WriteLine("[!] New E-mail Received.");
            switch (TypeInformation.GetTypeName(item))
            {
                case "MailItem":
                    {
                        MailItem itemCur = (MailItem)item;
                        Console.WriteLine(itemCur.Subject);
                        Console.WriteLine("=============================");
                        Console.WriteLine(itemCur.Body);
                        Console.WriteLine();
                        break;
                    }
                case "MeetingItem":
                    {
                        MeetingItem itemCur = (MeetingItem)item;
                        Console.WriteLine(itemCur.Subject);
                        Console.WriteLine("=============================");
                        Console.WriteLine(itemCur.Body);
                        Console.WriteLine();
                        break;
                    }
                default:
                    break;
            }
        }
        static Items GetInboxItems(OlDefaultFolders folder)
        {
            Application outlookApplication = new Application();
            NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
            MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(folder);
            return inboxFolder.Items;
        }
        public static class TypeInformation
        {
            public static string GetTypeName(object comObject)
            {
                var dispatch = comObject as IDispatch;

                if (dispatch == null)
                {
                    return null;
                }

                var pTypeInfo = dispatch.GetTypeInfo(0, 1033);

                string pBstrName;
                string pBstrDocString;
                int pdwHelpContext;
                string pBstrHelpFile;
                pTypeInfo.GetDocumentation(
                    -1,
                    out pBstrName,
                    out pBstrDocString,
                    out pdwHelpContext,
                    out pBstrHelpFile);

                string str = pBstrName;
                if (str[0] == 95)
                {
                    // remove leading '_'
                    str = str.Substring(1);
                }

                return str;
            }

            [ComImport]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            [Guid("00020400-0000-0000-C000-000000000046")]
            private interface IDispatch
            {
                int GetTypeInfoCount();

                [return: MarshalAs(UnmanagedType.Interface)]
                ITypeInfo GetTypeInfo(
                    [In, MarshalAs(UnmanagedType.U4)] int iTInfo,
                    [In, MarshalAs(UnmanagedType.U4)] int lcid);

                void GetIDsOfNames(
                    [In] ref Guid riid,
                    [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames,
                    [In, MarshalAs(UnmanagedType.U4)] int cNames,
                    [In, MarshalAs(UnmanagedType.U4)] int lcid,
                    [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
            }
        }
        //method to send email to outlook
        //https://www.codeproject.com/Tips/165548/Csharp-Code-Snippet-to-Send-an-Email-with-Attachme
        static void SendEmail(string recipient, string body, string subject)
        {
            try
            {
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = body;
                //Add an attachment.
                //String sDisplayName = "MyAttachment";
                int iPosition = (int)oMsg.Body.Length + 1;
                //int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                //Outlook.Attachment oAttach = oMsg.Attachments.Add
                //                             (@"C:\\fileName.jpg", iAttachType, iPosition, sDisplayName);
                //Subject line
                oMsg.Subject = subject;
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.

                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(recipient);
                oRecip.Resolve();
                // Send.
                oMsg.Send();
                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
            }//end of try block
            catch (System.Exception ex)
            {
            }//end of catch
        }
    }
}
