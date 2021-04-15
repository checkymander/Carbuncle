using Microsoft.Office.Interop.Outlook;
using System;
using Carbuncle.Helpers;
using System.Text.RegularExpressions;
using System.IO;

namespace Carbuncle.Commands
{
    public class MailSearcher
    {
        bool force { get; set; }
        public MailSearcher(bool force = false)
        {
            this.force = force;
        }
        public object ReadEmailByID(string guid, OlDefaultFolders folder)
        {
            Application outlookApplication = new Application();
            NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
            MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(folder);
            var item = outlookNamespace.GetItemFromID(guid);

            return item;
        }
        public void ReadEmailByNumber(int number)
        {
            Console.WriteLine("[+] Reading e-mail number: {0}", number);
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            try
            {
                var item = mailItems[number];
                if (item is MailItem mailItem)
                {
                    Common.DisplayMailItem(mailItem);
                }

                if (item is MeetingItem meetingItem)
                {
                    Common.DisplayMeetingItem(meetingItem);

                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        public void ReadEmailBySubject(string subject)
        {
            MailSearcher ms = new MailSearcher();
            Items mailItems = ms.GetInboxItems(OlDefaultFolders.olFolderInbox);
            Console.WriteLine("[+] Searching for e-mails, with the subject: {0}", subject);
            try
            {
                foreach (var item in mailItems)
                {
                    if (item is MailItem mailItem)
                    {
                        string body = "";
                        if (!String.IsNullOrEmpty(mailItem.Body))
                            body = mailItem.Body;
                        if (mailItem.Subject.Contains(subject))
                        {
                            Console.WriteLine(body);
                        }
                    }

                    if (item is MeetingItem meetingItem)
                    {
                        string body = "";
                        if (!String.IsNullOrEmpty(meetingItem.Body))
                            body = meetingItem.Body;
                        if (meetingItem.Subject.Contains(subject))
                        {
                            Console.WriteLine(body);
                        }
                    }
                }
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        public void SearchBySenderName(string name)
        {
            MailSearcher ms = new MailSearcher();

            Items mailItems = ms.GetInboxItems(OlDefaultFolders.olFolderInbox);
            Console.WriteLine("[+] Searching for e-mails from: {0}", name);
            foreach (var item in mailItems)
            {
                try
                {


                    if (item is MailItem mailItem)
                    {
                        if (mailItem.SenderName.ToLower().Contains(name.ToLower()))
                            Common.DisplayMailItem(mailItem);
                    }

                    if (item is MeetingItem meetingItem)
                    {
                        if (meetingItem.SenderEmailAddress.ToLower().Contains(name.ToLower()))
                            Common.DisplayMeetingItem(meetingItem);

                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        public void SearchBySubject(string subject)
        {
            SearchBySubjectRegex($"({subject})");
        }
        public void SearchBySubjectRegex(string regex)
        {
            Console.WriteLine("Searching by Subject Regex: " + regex);
            Regex r = new Regex(regex);
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            foreach (var item in mailItems)
            {
                try
                {
                    if (item is MailItem mailItem)
                    {
                        if (Regex.Match(mailItem.Subject, regex).Success)
                        {
                            Common.DisplayMailItem(mailItem);
                        }
                    }
                    else if (item is MeetingItem meetingItem)
                    {
                        if (Regex.Match(meetingItem.SenderEmailAddress, regex).Success)
                        {
                            Common.DisplayMeetingItem(meetingItem);
                        }
                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        public void SearchByAddress(string email)
        {
            /**
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            Console.WriteLine("[+] Searching for e-mails from: {0}", email);
            foreach (var item in mailItems)
            {
                try
                {
                    if (item is MailItem mailItem)
                    {
                        if (mailItem.SenderEmailAddress.ToLower().Contains(email.ToLower()))
                            Common.DisplayMailItem(mailItem);
                    }

                    if (item is MeetingItem meetingItem)
                    {
                        if (meetingItem.SenderEmailAddress.ToLower().Contains(email.ToLower()))
                            Common.DisplayMeetingItem(meetingItem);

                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            **/
            SearchByAddressRegex($"({email})");
        }
        public void SearchByAddressRegex(string regex)
        {
            //Can probably modify "SearchByKeyword" to make use of this function for less code.
            Regex r = new Regex(regex);
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            foreach (var item in mailItems)
            {
                try
                {
                    if (item is MailItem mailItem)
                    {
                        if (!String.IsNullOrEmpty(mailItem.Subject) && Regex.Match(mailItem.Subject, regex).Success)
                        {
                            Common.DisplayMailItem(mailItem);
                        }
                    }
                    else if (item is MeetingItem meetingItem)
                    {
                        if (!String.IsNullOrEmpty(meetingItem.Subject) && Regex.Match(meetingItem.Subject, regex).Success)
                        {
                            Common.DisplayMeetingItem(meetingItem);
                        }
                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        public void SearchByContent(string[] content)
        {
            /**
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            Console.WriteLine("[+] Searching for e-mails that contain the keyword(s): {0}", content);
            foreach (var item in mailItems)
            {
                try
                {
                    if (item is MailItem mailItem)
                    {
                        string body = "";
                        if (!String.IsNullOrEmpty(mailItem.Body))
                            body = mailItem.Body;

                        if (content == "" || body.ToLower().Contains(content.ToLower()) || mailItem.Subject.ToLower().Contains(content.ToLower()))
                            Common.DisplayMailItem(mailItem);
                    }

                    if (item is MeetingItem meetingItem)
                    {
                        string body = "";
                        if (!String.IsNullOrEmpty(meetingItem.Body))
                            body = meetingItem.Body;
                        if (content == "" || body.ToLower().Contains(content.ToLower()) || meetingItem.Subject.ToLower().Contains(content.ToLower()))
                            Common.DisplayMeetingItem(meetingItem);
                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            
            }
            **/
            string regex = "";
            foreach(var s in content)
            {
                regex += $"({s})";
            }
            SearchByContentRegex(regex);
        }
        public void SearchByContentRegex(string regex)
        {
            Regex r = new Regex(regex);
            Items mailItems = GetInboxItems(OlDefaultFolders.olFolderInbox);
            foreach (var item in mailItems)
            {
                try
                {
                    if (item is MailItem mailItem)
                    {
                        if (!String.IsNullOrEmpty(mailItem.Body) && Regex.Match(mailItem.Body, regex).Success)
                        {
                            Common.DisplayMailItem(mailItem);
                        }
                    }
                    else if (item is MeetingItem meetingItem)
                    {
                        if (!String.IsNullOrEmpty(meetingItem.Body) && Regex.Match(meetingItem.Body, regex).Success)
                        {
                            Common.DisplayMeetingItem(meetingItem);
                        }
                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }     
        public Items GetInboxItems(OlDefaultFolders folder)
        {
            Application outlookApplication = new Application();
            NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
            MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(folder);
            return inboxFolder.Items;
        }
        public void GetAll(bool display = false)
        {
            MailSearcher ms = new MailSearcher();
            Items mailItems = ms.GetInboxItems(OlDefaultFolders.olFolderInbox);
            if (mailItems.Count > 200 && !force)
            {
                Console.WriteLine("[!] Warning: You are about to display the information of over 200 e-mail subjects. Are you sure you don't want to search by keyword or name? Use /force to bypass this warning.\r\n[!] Current Count: {0}", mailItems.Count);
                return;
            }
            Console.WriteLine("[+] Getting all e-mail items");
            foreach (var item in mailItems)
            {
                try
                {
                    if (item is MailItem mailItem)
                    {
                        Common.DisplayMailItem(mailItem);
                    }

                    if (item is MeetingItem meetingItem)
                    {
                        Common.DisplayMeetingItem(meetingItem);

                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
    }
}
