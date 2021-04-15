using Carbuncle.Helpers;
using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Carbuncle.Commands
{
    public class AttachmentSearcher
    {
        bool download { get; set; }
        string downloadFolder { get; set; }
        public AttachmentSearcher(bool download = false, string downloadFolder = "")
        {
            this.download = download;
            this.downloadFolder = downloadFolder;
        }
        public void GetAttachmentsByID(string ID, OlDefaultFolders folder)
        {
            Application outlookApplication = new Application();
            NameSpace outlookNamespace = outlookApplication.GetNamespace("MAPI");
            MAPIFolder inboxFolder = outlookNamespace.GetDefaultFolder(folder);
            var item = outlookNamespace.GetItemFromID(ID);
            if (item is MailItem mailItem)
            {
                if (mailItem.Attachments.Count > 0)
                {
                    if (download)
                    {
                        foreach (Attachment attachment in mailItem.Attachments)
                        {
                            attachment.SaveAsFile(downloadFolder.TrimEnd('\\') + "\\" + attachment.FileName);
                        }
                    }
                    else
                    {
                        Common.DisplayMailItem(mailItem);
                    }
                }
            }
            else if (item is MeetingItem meetingItem)
            {
                if (meetingItem.Attachments.Count > 0)
                {
                    if (download)
                    {
                        foreach (Attachment attachment in meetingItem.Attachments)
                        {
                            attachment.SaveAsFile(downloadFolder.TrimEnd('\\') + "\\" + attachment.FileName);
                        }
                    }
                    else
                    {
                        Common.DisplayMeetingItem(meetingItem);
                    }
                }
            }
        }
        public void GetAttachmentsByKeyword(string keyword)
        {
            GetAttachmentsByRegex($"({keyword})");
        }
        public void GetAttachmentsByRegex(string regex)
        {
            MailSearcher ms = new MailSearcher();
            Items mailItems = ms.GetInboxItems(OlDefaultFolders.olFolderInbox);
            foreach (var item in ms.GetInboxItems(OlDefaultFolders.olFolderInbox))
            {
                try
                {
                    if (item is MailItem mailItem)
                    {
                        if (mailItem.Attachments.Count > 0)
                        {
                            foreach (Attachment attachment in mailItem.Attachments)
                            {
                                if (Regex.Match(attachment.FileName, regex).Success)
                                {
                                    if (download)
                                    {
                                        attachment.SaveAsFile(downloadFolder.TrimEnd('\\') + "\\" + attachment.FileName);
                                    }
                                    else
                                    {
                                        Common.DisplayMailItem(mailItem);
                                    }
                                }

                            }
                        }
                    }
                    else if (item is MeetingItem meetingItem)
                    {
                        if (meetingItem.Attachments.Count > 0)
                        {
                            foreach (Attachment attachment in meetingItem.Attachments)
                            {
                                if(Regex.Match(attachment.FileName, regex).Success)
                                {
                                    if (download)
                                    {
                                        attachment.SaveAsFile(downloadFolder.TrimEnd('\\') + "\\" + attachment.FileName);
                                    }
                                    else
                                    {
                                        Common.DisplayMeetingItem(meetingItem);
                                    }
                                }

                            }
                        }
                    }
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
        }
        public void GetAllAttachments()
        {
            Console.WriteLine("Getting All Attachments\r\nDownload = " + download);
            Console.WriteLine(downloadFolder);
            MailSearcher ms = new MailSearcher();
            //Items mailItems = ms.GetInboxItems(OlDefaultFolders.olFolderInbox);

            foreach (var item in ms.GetInboxItems(OlDefaultFolders.olFolderInbox))
            {
                try
                {
                    if (item is MailItem mailItem)
                    {
                        if (mailItem.Attachments.Count > 0)
                        {
                            if (download)
                            {
                                foreach (Attachment attachment in mailItem.Attachments)
                                {
                                    attachment.SaveAsFile(downloadFolder.TrimEnd('\\') + "\\" + attachment.FileName);
                                }
                            }
                            else
                            {
                                Common.DisplayMailItem(mailItem);
                            }
                        }
                    }
                    else if (item is MeetingItem meetingItem)
                    {
                        if (meetingItem.Attachments.Count > 0)
                        {
                            if (download)
                            {
                                foreach (Attachment attachment in meetingItem.Attachments)
                                {
                                    attachment.SaveAsFile(downloadFolder.TrimEnd('\\') + "\\" + attachment.FileName);
                                }
                            }
                            else
                            {
                                Common.DisplayMeetingItem(meetingItem);
                            }
                        }
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
