using Microsoft.Office.Interop.Outlook;
using System;

namespace Carbuncle.Helpers
{
    public class Common
    {
        public static bool display = false;
        public static bool force = false;
        public static void DisplayMailItem(MailItem item)
        {
            Console.WriteLine("[Sender] {0} - ({1})", item.SenderName, item.SenderEmailAddress);
            Console.WriteLine("[Subject] " + item.Subject);
            Console.WriteLine("[ID] " + item.EntryID);
            if(item.Attachments.Count > 0)
            {
                Console.Write("[Attachments]");
                foreach(Attachment attach in item.Attachments)
                {
                    Console.Write(" " + attach.FileName);
                }
                Console.WriteLine();
            }
                
            if (Common.display)
                Console.WriteLine("[Body] " + item.Body);

            Console.WriteLine();
        }
        public static void DisplayMeetingItem(MeetingItem item)
        {
            Console.WriteLine("[Sender] {0} - ({1})", item.SenderName, item.SenderEmailAddress);
            Console.WriteLine("[Subject] " + item.Subject);
            Console.WriteLine("[ID] " + item.EntryID);
            if (item.Attachments.Count > 0)
            {
                Console.Write("[Attachments]");
                foreach (Attachment attach in item.Attachments)
                {
                    Console.Write(" " + attach.FileName);
                }
                Console.WriteLine();
            }
            if (Common.display)
                Console.WriteLine("[Body] " + item.Body);
            Console.WriteLine();
        }
        public static void PrintHelp()
        {

            //carbuncle.exe searchmail /body [/regex:"blahblahblah"] [/content:"blahblahblah"]
            //carbuncle.exe searchmail /senderaddress [/regex:"blahblahblah"] [/address:"checkymander@protonmail.com"]
            //carbuncle.exe searchmail /subject [/regex:"blahblahblah"] [/content:"blahblahblah"]
            //carbuncle.exe searchmail /attachment [/regex:"blahblahblah"] [/name:"blahblahblah"] [/download] [/downloadpath:"C:\\users\\checkymander\\Documents\\"]
            //carbuncle.exe read /entryid:00000000ABF08F38F774EF44BD800D54DA6135740700438C90E5F1E27549A26DD4C4CE7C884C0069B971A0EB00007E3487BFEF2F834F93D188D339E4EA4E00003BA5A49B0000
            //Can you reference an e-mail by its ID?
            string helptext = @"Carbuncle Usage:
carbuncle.exe searchmail [/senderaddress:test@email.com] [/sendername:""Mander, Checky""] [/keyword:P@ssw0rd] [/display]
carbuncle.exe attachments [TODO]
carbuncle.exe read [/subject:""Important E-mail""] [/number:10]
carbuncle.exe send /body:""This is an important e-mail body""  /subject:""Important e-mail'"" /recipients:""test@gmail.com,test2@gmail.com"" [/attachment:""C:\users\checkymander\pictures\picture.jpg""] [/attachmentname:picture.jpg]
carbuncle.exe monitor [/display]";

            Console.WriteLine(helptext);
        }
    }
}
