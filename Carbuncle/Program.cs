using Microsoft.Office.Interop.Outlook;
using System;
using Exception = System.Exception;
using System.IO;


namespace Carbuncle
{
    class Program
    {
        static bool display = false;
        static void Main(string[] args)
        {
            /**
             * ToDo:
]            * Has Attachment
             * */

            var parsed = ArgumentParser.Parse(args);
            if (!parsed.ParsedOk)
            {
                return;
            }
            string action = args.Length != 0 ? args[0] : "";
            
            if (parsed.Arguments.ContainsKey("display"))
            {
                Console.WriteLine("[+] Setting to display e-mails");
                display = true;
            }

            switch (action.ToLower())
            {
                case "read":
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
                case "enum":
                    if (parsed.Arguments.ContainsKey("keyword"))
                    {
                        SearchByKeyword(parsed.Arguments["keyword"]);
                    }
                    else if (parsed.Arguments.ContainsKey("email"))
                    {
                        SearchByEmail(parsed.Arguments["email"]);
                    }
                    else if (parsed.Arguments.ContainsKey("name"))
                    {
                        SearchByName(parsed.Arguments["name"]);
                    }
                    else
                    {
                        GetAll();
                    }
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
        }
        static void PrintHelp()
        {
            string helptext = @"Carbuncle Usage:
carbuncle.exe enum [/email:test@email.com] [/name:""Mander, Checky""] [/keyword:P@ssw0rd] [/display]
carbuncle.exe read [/subject:""Important E-mail""] [/number:10]
carbuncle.exe send /body:""This is an important e-mail body""  /subject:""Important e-mail'"" /recipients:""test@gmail.com,test2@gmail.com"" [/attachment:""C:\users\checkymander\pictures\picture.jpg""] [/attachmentname:picture.jpg]
carbuncle.exe monitor [/display]";
            
            Console.WriteLine(helptext);
        }
        static Items GetInboxItems(OlDefaultFolders folder)
        {
            Application outlookApplication = new Application();
            NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
            MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(folder);
            return inboxFolder.Items;
        }
        static void ReadEmail(string Subject)
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            Console.WriteLine("[+] Searching for e-mails, with the subject: {0}",Subject);
            try
            {
                foreach (var item in mailItems)
                {
                    switch (TypeInformation.GetTypeName(item))
                    {
                        case "MailItem":
                            {
                                string body = "";
                                MailItem itemCur = (MailItem)item;
                                if (!String.IsNullOrEmpty(itemCur.Body))
                                    body = itemCur.Body;

                                if (itemCur.Subject.Contains(Subject))
                                {
                                    Console.WriteLine(body);
                                }
                                break;
                            }
                        case "MeetingItem":
                            {
                                string body = "";
                                MeetingItem itemCur = (MeetingItem)item;
                                if (!String.IsNullOrEmpty(itemCur.Body))
                                    body = itemCur.Body;

                                if (itemCur.Subject.Contains(Subject))
                                {
                                    Console.WriteLine(body);
                                }
                                break;
                            }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        static void ReadEmail(int number)
        {
            Console.WriteLine("[+] Reading e-mail number: {0}", number);
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            try
            {
                var item = mailItems[number];

                switch (TypeInformation.GetTypeName(item))
                {
                    case "MailItem":
                        {
                            MailItem itemCur = (MailItem)item;
                            DisplayMailItem(itemCur);

                            break;
                        }
                    case "MeetingItem":
                        {
                            MeetingItem itemCur = (MeetingItem)item;
                            DisplayMeetingItem(itemCur);

                            break;
                        }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        static void SearchByName(string Name)
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            Console.WriteLine("[+] Searching for e-mails from: {0}", Name);
            foreach (var item in mailItems)
            {
                try
                {
                    switch (TypeInformation.GetTypeName(item))
                    {
                        case "MailItem":
                            {
                                MailItem itemCur = (MailItem)item;
                                if (itemCur.SenderName.ToLower().Contains(Name.ToLower()))
                                    DisplayMailItem(itemCur);

                                break;
                            }
                        case "MeetingItem":
                            {
                                MeetingItem itemCur = (MeetingItem)item;
                                if (itemCur.SenderEmailAddress.ToLower().Contains(Name.ToLower()))
                                    DisplayMeetingItem(itemCur);

                                break;
                            }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        static void SearchByEmail(string Email)
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            Console.WriteLine("[+] Searching for e-mails from: {0}", Email);
            foreach (var item in mailItems)
            {
                try
                {
                    switch (TypeInformation.GetTypeName(item))
                    {
                        case "MailItem":
                            {
                                MailItem itemCur = (MailItem)item;
                                if (itemCur.SenderEmailAddress.ToLower().Contains(Email.ToLower()))
                                    DisplayMailItem(itemCur);

                                break;
                            }
                        case "MeetingItem":
                            {
                                MeetingItem itemCur = (MeetingItem)item;
                                if (itemCur.SenderEmailAddress.ToLower().Contains(Email.ToLower()))
                                    DisplayMeetingItem(itemCur);

                                break;
                            }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        static void SearchByKeyword(string keyword)
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            Console.WriteLine("[+] Searching for e-mails that contain the keyword(s): {0}", keyword);

            foreach (var item in mailItems)
            {
                try
                {
                    switch (TypeInformation.GetTypeName(item))
                    {
                        case "MailItem":
                            {
                                string body = "";
                                MailItem itemCur = (MailItem)item;
                                if (!String.IsNullOrEmpty(itemCur.Body))
                                    body = itemCur.Body;

                                if (keyword == "" || body.ToLower().Contains(keyword.ToLower()) || itemCur.Subject.ToLower().Contains(keyword.ToLower()))
                                    DisplayMailItem(itemCur);

                                break;
                            }
                        case "MeetingItem":
                            {
                                string body = "";
                                MeetingItem itemCur = (MeetingItem)item;
                                if (!String.IsNullOrEmpty(itemCur.Body))
                                    body = itemCur.Body;
                                if (keyword == "" || body.ToLower().Contains(keyword.ToLower()) || itemCur.Subject.ToLower().Contains(keyword.ToLower()))
                                    DisplayMeetingItem(itemCur);

                                break;
                            }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }

        static void GetAll()
        {
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            Console.WriteLine("[+] Getting all e-mail items");
            foreach (var item in mailItems)
            {
                try
                {
                    switch (TypeInformation.GetTypeName(item))
                    {
                        case "MailItem":
                            {
                            MailItem itemCur = (MailItem)item;
                            DisplayMailItem(itemCur);
                            break;
                            }
                        case "MeetingItem":
                            {
                            MeetingItem itemCur = (MeetingItem)item;
                            DisplayMeetingItem(itemCur);
                            break; 
                            }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        static void MonitorEmail()
        {
            Console.WriteLine("[+] Starting e-mail monitoring...");
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            mailItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(NewEmailEvent);
            Console.WriteLine("[+] Started, press Ctrl+Z to exit");
        }
        static void NewEmailEvent(object item)
        {
            switch (TypeInformation.GetTypeName(item))
            {
                case "MailItem":
                    {
                        MailItem itemCur = (MailItem)item;
                        DisplayMailItem(itemCur);
                        break;
                    }
                case "MeetingItem":
                    {
                        MeetingItem itemCur = (MeetingItem)item;
                        DisplayMeetingItem(itemCur);
                        break;
                    }
                default:
                    break;
            }
        }
        static void SendEmail(string[] recipients, string body, string subject)
        {
            Console.WriteLine("[+] Sending an e-mail.\r\nRecipients: {0}\r\nSubject: {1}\r\nBody: {2}", String.Join(",", recipients), subject, body);
            try
            {
                Application outlookApplication = new Application();
                MailItem msg = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
                msg.HTMLBody = body;
                msg.Subject = subject;
                foreach(var recipient in recipients)
                {
                    Recipients recips = msg.Recipients;
                    Recipient recip = recips.Add(recipient);
                    recip.Resolve();

                }
                msg.Send();
                Console.WriteLine("[+] Message Sent");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        static void SendEmail(string[] recipients, string body, string subject, string attachment, string attachmentname)
        {
            Console.WriteLine("[+] Sending an e-mail.\r\nRecipients: {0}\r\nSubject: {1}\r\nAttachment Path: {2}\r\nAttachment Name: {3}\r\nBody: {4}",String.Join(",",recipients),subject,attachment,attachmentname,body);
            try
            {
                Application outlookApplication = new Application();
                MailItem msg = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
                msg.HTMLBody = body;
                int pos = msg.Body.Length + 1;
                int attType = (int)OlAttachmentType.olByValue;
                Attachment attach = msg.Attachments.Add(attachment, attType, pos, attachmentname);
                msg.Subject = subject;
                foreach (var recipient in recipients)
                {
                    Recipients recips = msg.Recipients;
                    Recipient recip = recips.Add(recipient);
                    recip.Resolve();

                }
                msg.Send();
                Console.WriteLine("[+] Message Sent");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        static void DisplayMailItem(MailItem item)
        {
            Console.WriteLine("[Sender] {0} - ({1})", item.SenderName, item.SenderEmailAddress);
            Console.WriteLine("[Subject] " + item.Subject);
            if (display)
                Console.WriteLine("[Body] " + item.Body);
            Console.WriteLine();
        }
        static void DisplayMeetingItem(MeetingItem item)
        {
            Console.WriteLine("[Sender] {0} - ({1})", item.SenderName, item.SenderEmailAddress);
            Console.WriteLine("[Subject] " + item.Subject);
            if (display)
                Console.WriteLine("[Body] " + item.Body);
            Console.WriteLine();
        }
    }
}
