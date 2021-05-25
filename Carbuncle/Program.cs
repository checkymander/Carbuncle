using System;
using System.IO;

using Carbuncle.Helpers;
using Carbuncle.Commands;
using Microsoft.Office.Interop.Outlook;

namespace Carbuncle
{
    class Program
    {
        static void Main(string[] args)
        {
            var parsed = ArgumentParser.Parse(args);
            if (!parsed.ParsedOk)
            {
                return;
            }
            string action = args.Length != 0 ? args[0] : "";
            
            if (parsed.Arguments.ContainsKey("display"))
            {
                Console.WriteLine("[+] Setting to display e-mails");
                Common.display = true;
            }
            if (parsed.Arguments.ContainsKey("force"))
            {
                Console.WriteLine("[+] Enabling force");
                Common.force = true;
            }
            MailSearcher ms = new MailSearcher(force:Common.force);
            AttachmentSearcher at = new AttachmentSearcher();
            switch (action.ToLower())
            {
                case "read":
                    Common.display = true;
                    if (parsed.Arguments.ContainsKey("number"))
                    {
                        ms.ReadEmailByNumber(int.Parse(parsed.Arguments["number"]));
                    }
                    else if (parsed.Arguments.ContainsKey("subject"))
                    {
                        ms.ReadEmailBySubject(parsed.Arguments["subject"]);
                    }
                    else if (parsed.Arguments.ContainsKey("entryid"))
                    {
                        var item = ms.ReadEmailByID(parsed.Arguments["entryid"], OlDefaultFolders.olFolderInbox);
                        if (item is MailItem mailItem)
                        {
                            Common.DisplayMailItem(mailItem);
                        }
                        else if (item is MeetingItem meetingItem)
                        {
                            Common.DisplayMeetingItem(meetingItem);
                        }
                    }
                    else
                    {
                        Common.PrintHelp();
                    }
                    break;
                case "searchmail":
                    string searchMethod;
                    try
                    {
                        searchMethod = args[1].TrimStart('/');
                    }
                    catch
                    {
                        searchMethod = "all";
                    }
                    switch(searchMethod.ToLower()){
                        case "body":
                            if (parsed.Arguments.ContainsKey("regex"))
                            {
                                ms.SearchByContentRegex(parsed.Arguments["regex"]);
                            }
                            else if (parsed.Arguments.ContainsKey("content"))
                            {
                                string[] keywords = parsed.Arguments["content"].Split(',');
                                ms.SearchByContent(keywords);
                            }
                            break;
                        case "senderaddress":
                            if (parsed.Arguments.ContainsKey("regex"))
                            {
                                ms.SearchByAddressRegex(parsed.Arguments["regex"]);
                            }
                            else if (parsed.Arguments.ContainsKey("address"))
                            {
                                ms.SearchByAddress(parsed.Arguments["address"]);
                            }
                            break;
                        case "subject":
                            if (parsed.Arguments.ContainsKey("regex"))
                            {
                                ms.SearchBySubjectRegex(parsed.Arguments["regex"]);
                            }
                            else if (parsed.Arguments.ContainsKey("content"))
                            {
                                ms.SearchBySubject(parsed.Arguments["content"]);
                            }
                            break;
                        case "attachment":
                            if (parsed.Arguments.ContainsKey("regex"))
                            {
                                if (parsed.Arguments.ContainsKey("download")){
                                    if (parsed.Arguments.ContainsKey("downloadpath"))
                                    {
                                        at = new AttachmentSearcher(true, parsed.Arguments["downloadpath"]);
                                        at.GetAttachmentsByRegex(parsed.Arguments["regex"]);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Missing download path parameter!");
                                        Common.PrintHelp();
                                    }
                                }
                                else
                                {
                                    at.GetAttachmentsByRegex(parsed.Arguments["regex"]);
                                }
                            }
                            else if (parsed.Arguments.ContainsKey("name"))
                            {
                                if (parsed.Arguments.ContainsKey("download"))
                                {
                                    if (parsed.Arguments.ContainsKey("downloadpath"))
                                    {
                                        at = new AttachmentSearcher(true, parsed.Arguments["downloadpath"]);
                                        at.GetAttachmentsByKeyword(parsed.Arguments["name"]);
                                    }
                                    else
                                    {
                                        Console.WriteLine("Missing download path parameter!");
                                        Common.PrintHelp();
                                    }
                                }
                                else
                                {
                                    at.GetAttachmentsByKeyword(parsed.Arguments["name"]);
                                }
                            }
                            break;
                        case "all":
                            ms.GetAll();
                            break;
                        default:
                            ms.GetAll();
                            break;

                    }
                    break;
                case "monitor":
                    MailMonitor mm = new MailMonitor();
                    if (parsed.Arguments.ContainsKey("regex"))
                    {
                        mm.Start(parsed.Arguments["regex"]);
                    }
                    else
                    {
                        mm.Start();
                    }
                    while (true)
                    {

                    }
                    break;
                case "send":
                    if (parsed.Arguments.ContainsKey("recipients") && parsed.Arguments.ContainsKey("subject") && parsed.Arguments.ContainsKey("body"))
                    {
                        MailSender sender = new MailSender();
                        if (parsed.Arguments.ContainsKey("attachment"))
                        {
                            string AttachmentName;
                            if (parsed.Arguments.ContainsKey("attachmentname"))
                                AttachmentName = parsed.Arguments["attachmentname"];
                            else
                                AttachmentName = Path.GetFileNameWithoutExtension(parsed.Arguments["attachment"]);
                            sender.SendEmail(parsed.Arguments["recipients"].Split(','), parsed.Arguments["body"], parsed.Arguments["subject"], parsed.Arguments["attachment"], AttachmentName);
                        }
                        else
                        {
                            sender.SendEmail(parsed.Arguments["recipients"].Split(','), parsed.Arguments["body"], parsed.Arguments["subject"]);
                        }
                    }          
                    break;
                case "attachments":
                    if (!parsed.Arguments.ContainsKey("downloadpath"))
                    {
                        Console.WriteLine("Missing downloadpath parameter!");
                        Common.PrintHelp();
                        break;
                    }
                    at = new AttachmentSearcher(download: true, downloadFolder: parsed.Arguments["downloadpath"]);

                    if (parsed.Arguments.ContainsKey("all"))
                    {
                        at.GetAllAttachments();
                    }
                    else if (parsed.Arguments.ContainsKey("entryid"))
                    {
                        at.GetAttachmentsByID(parsed.Arguments["entryid"], OlDefaultFolders.olFolderInbox);
                    }
                    break;
                default:
                    Common.PrintHelp();
                    break;
            }
            
            Console.WriteLine("Done.");
        }

    }
}
