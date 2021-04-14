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
            if (Common.display)
                Console.WriteLine("[Body] " + item.Body);
            Console.WriteLine();
        }
        public static void DisplayMeetingItem(MeetingItem item)
        {
            Console.WriteLine("[Sender] {0} - ({1})", item.SenderName, item.SenderEmailAddress);
            Console.WriteLine("[Subject] " + item.Subject);
            if (Common.display)
                Console.WriteLine("[Body] " + item.Body);
            Console.WriteLine();
        }
        public static void PrintHelp()
        {

            //carbuncle.exe searchmail /content [/regex:"blahblahblah"] [/keyword:"blahblahblah"]
            //carbuncle.exe searchmail /senderaddress [/regex:"blahblahblah"] [/address:"checkymander@protonmail.com"]
            //carbuncle.exe searchmail /subject [/regex:"blahblahblah"] [/subject:"blahblahblah"]
            //carbuncle.exe searchmail /attachment [/regex:"blahblahblah"] [/name:"blahblahblah"]
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
