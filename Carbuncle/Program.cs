using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;


namespace Carbuncle
{
    class Program
    {
        static bool display = false;
        static bool verbose = false;
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
            
            if (parsed.Arguments.ContainsKey("display"))
            {
                Console.WriteLine("Setting Display to True");
                display = true;
            }
            if (parsed.Arguments.ContainsKey("verbose"))
            {
                verbose = true;
            }

            switch (action.ToLower())
            {
                case "read":
                    //Read from Inbox Number or Subject
                    if (parsed.Arguments.ContainsKey("number"))
                    {
                        ReadEmail(int.Parse(parsed.Arguments["number"]));
                    }
                    else if (parsed.Arguments.ContainsKey("subject"))
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
                    {
                        if (parsed.Arguments.ContainsKey("recipients") && parsed.Arguments.ContainsKey("subject") && parsed.Arguments.ContainsKey("body"))
                        {
                            if (parsed.Arguments.ContainsKey("attachment"))
                            {
                                string AttachmentName;
                                if (parsed.Arguments.ContainsKey("attachmentname"))
                                    AttachmentName = parsed.Arguments["attachmentname"];
                                else
                                    AttachmentName = Path.GetFileNameWithoutExtension(parsed.Arguments["attachment"]);
                                
                                SendEmail(parsed.Arguments["recipients"].Split(','), parsed.Arguments["body"], parsed.Arguments["subject"], parsed.Arguments["attachment"], AttachmentName);
                            }
                            else
                            {
                                SendEmail(parsed.Arguments["recipients"].Split(','), parsed.Arguments["body"], parsed.Arguments["subject"]);
                            }
                        }
                            
                        break;
                    }
                default:
                    PrintHelp();
                    Console.ReadKey();
                    break;
            }
            Console.WriteLine("Press Any Key To Exit");
            Console.ReadKey();
        }
        static void PrintHelp()
        {
            Console.WriteLine("Carbuncle Usage:\r\ncarbuncle.exe enum\r\ncarbuncle.exe search / keyword:\"password\"\r\ncarbuncle.exe send / body:\"Hello World\" / subject:\"Subject E-mail\" / recipient:\"test@email.com\"\r\ncarbuncle.exe read / subject:\"Subject of E-mail\"\r\ncarbuncle.exe read / number:\"15\"\r\ncarbuncle.exe monitor");
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
            try
            {
                MailItem item = (MailItem)mailItems[number];
                Console.WriteLine(item.Subject);
                Console.WriteLine(item.Body);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        static void SearchAll(string keyword)
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);

                foreach (var item in mailItems)
                {
                    try
                    {
                        switch (TypeInformation.GetTypeName(item))
                        {
                            case "MailItem":
                                {
                                    MailItem itemCur = (MailItem)item;
                                    if (keyword == "" || itemCur.Body.ToLower().Contains(keyword.ToLower()) || itemCur.Body.ToLower().Contains(keyword.ToLower()))
                                    {
                                        Console.WriteLine(itemCur.Subject);
                                        if (display)
                                        {
                                            Console.WriteLine(itemCur.Body);
                                            Console.WriteLine();
                                        }
                                    }
                                    break;
                                }
                            case "MeetingItem":
                                {
                                    MeetingItem itemCur = (MeetingItem)item;
                                    if (keyword == "" || itemCur.Body.ToLower().Contains(keyword.ToLower()) || itemCur.Body.ToLower().Contains(keyword.ToLower()))
                                    {
                                        Console.WriteLine(itemCur.Subject);
                                        if (display)
                                        {
                                            Console.WriteLine(itemCur.Body);
                                            Console.WriteLine();
                                        }

                                    }
                                    break;
                                }
                        }
                    }
                    catch
                    {

                    }
                }
        }
        static void SearchMeetings()
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            List<MailItem> ReceivedEmail = new List<MailItem>();
            foreach (var item in mailItems)
            {
                try
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
                catch
                {

                }
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
                        if (display)
                        {
                            Console.WriteLine("=============================");
                            Console.WriteLine(itemCur.Body);
                        }
                        Console.WriteLine();
                        break;
                    }
                case "MeetingItem":
                    {
                        MeetingItem itemCur = (MeetingItem)item;
                        Console.WriteLine(itemCur.Subject);
                        if (display)
                        {
                            Console.WriteLine("=============================");
                            Console.WriteLine(itemCur.Body);
                        }
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
        static void SendEmail(string[] recipients, string body, string subject)
        {
            try
            {
                Outlook.Application outlookApplication = new Outlook.Application();
                Outlook.MailItem msg = (Outlook.MailItem)outlookApplication.CreateItem(Outlook.OlItemType.olMailItem);
                msg.HTMLBody = body;
                msg.Subject = subject;
                foreach(var recipient in recipients)
                {
                    Outlook.Recipients recips = (Outlook.Recipients)msg.Recipients;
                    Outlook.Recipient recip = (Outlook.Recipient)recips.Add(recipient);
                    recip.Resolve();

                }
                msg.Send();
                Console.WriteLine("Message Sent");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        static void SendEmail(string[] recipients, string body, string subject, string attachment, string attachmentname)
        {
            try
            {
                Outlook.Application outlookApplication = new Outlook.Application();
                Outlook.MailItem msg = (Outlook.MailItem)outlookApplication.CreateItem(Outlook.OlItemType.olMailItem);
                msg.HTMLBody = body;
                int pos = (int)msg.Body.Length + 1;
                int attType = (int)OlAttachmentType.olByValue;
                Outlook.Attachment attach = msg.Attachments.Add(attachment, attType, pos, attachmentname);
                msg.Subject = subject;
                foreach (var recipient in recipients)
                {
                    Outlook.Recipients recips = (Outlook.Recipients)msg.Recipients;
                    Outlook.Recipient recip = (Outlook.Recipient)recips.Add(recipient);
                    recip.Resolve();

                }
                msg.Send();
                Console.WriteLine("Message Sent");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
