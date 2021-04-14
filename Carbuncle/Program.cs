using System;
using System.IO;

using Carbuncle.Helpers;
using Carbuncle.Commands;

namespace Carbuncle
{
    class Program
    {
        static void Main(string[] args)
        {
            MailSearcher ms = new MailSearcher();
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

            switch (action.ToLower())
            {
                case "read":
                    if (parsed.Arguments.ContainsKey("number"))
                    {
                        ms.ReadEmailByNumber(int.Parse(parsed.Arguments["number"]));
                    }
                    else if (parsed.Arguments.ContainsKey("subject"))
                    {
                        ms.ReadEmailBySubject(parsed.Arguments["subject"]);
                    }
                    else
                    {
                        Common.PrintHelp();
                    }
                    break;
                case "searchmail":
                    //Step 1.) Determine if Search is by Display Name, E-mail Address, Body Content, AttachmentName, or Subject 
                    //Step 2.) Identify if they're searching by Regex or "Keyword"
                    var searchMethod = args[1].TrimStart('/');
                    switch(searchMethod.ToLower()){
                        case "content":
                            if (parsed.Arguments.ContainsKey("regex"))
                            {
                                ms.SearchByContent(parsed.Arguments["regex"]);
                            }
                            else if (parsed.Arguments.ContainsKey("keywords"))
                            {
                                string[] keywords = parsed.Arguments["keywords"].Split(',');
                                // Search by keywords
                                // Maybe change this?
                            }
                            break;
                        case "senderaddress":
                            if (parsed.Arguments.ContainsKey("regex"))
                            {
                                ms.SearchByAddress(parsed.Arguments["regex"]);
                            }
                            else if (parsed.Arguments.ContainsKey("address"))
                            {
                                ms.SearchByAddress(parsed.Arguments["senderaddress"]);
                            }
                            break;
                        case "subject":
                            if (parsed.Arguments.ContainsKey("regex"))
                            {
                                ms.SearchBySubjectRegex(parsed.Arguments["regex"]);
                            }
                            else if (parsed.Arguments.ContainsKey("subject"))
                            {
                                ms.SearchBySubject(parsed.Arguments["subject"]);
                            }
                            break;
                        case "attachment":
                            if (parsed.Arguments.ContainsKey("regex"))
                            {
                                //not implemented yet
                                throw new NotImplementedException("Method has not been implemented yet.") ;
                            }
                            else if (parsed.Arguments.ContainsKey("name"))
                            {
                                //not implemented yet
                                throw new NotImplementedException("Method has not been implemented yet.");
                            }
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
                        MailSender msend = new MailSender();
                        if (parsed.Arguments.ContainsKey("attachment"))
                        {
                            string AttachmentName;
                            if (parsed.Arguments.ContainsKey("attachmentname"))
                                AttachmentName = parsed.Arguments["attachmentname"];
                            else
                                AttachmentName = Path.GetFileNameWithoutExtension(parsed.Arguments["attachment"]);
                            msend.SendEmail(parsed.Arguments["recipients"].Split(','), parsed.Arguments["body"], parsed.Arguments["subject"], parsed.Arguments["attachment"], AttachmentName);
                        }
                        else
                        {
                            msend.SendEmail(parsed.Arguments["recipients"].Split(','), parsed.Arguments["body"], parsed.Arguments["subject"]);
                        }
                    }          
                    break;
                default:
                    Common.PrintHelp();
                    break;
            }
        }

    }
}
