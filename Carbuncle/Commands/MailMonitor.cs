using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using Carbuncle.Helpers;
using Exception = System.Exception;

namespace Carbuncle.Commands
{
    public class MailMonitor
    {
        public MailMonitor()
        {

        }
        private void NewEmailEvent(object item)
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
        public void Start()
        {
            MonitorEmail();
        }
        public void Start(string regex)
        {
            MonitorEmailRegex(regex);
        }
        
        private void MonitorEmail()
        {
            MailSearcher ms = new MailSearcher();
            Console.WriteLine("[+] Starting e-mail monitoring...");
            Items mailItems = ms.GetInboxItems(OlDefaultFolders.olFolderInbox);
            mailItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(NewEmailEvent);
            Console.WriteLine("[+] Started, press Ctrl+Z to exit");
        }
        private void MonitorEmailRegex(string regex)
        {
            MailSearcher ms = new MailSearcher();
            Console.WriteLine("[+] Starting e-mail monitoring...");
            Items mailItems = ms.GetInboxItems(OlDefaultFolders.olFolderInbox);
            mailItems.ItemAdd += new ItemsEvents_ItemAddEventHandler(NewEmailEvent);
            Console.WriteLine("[+] Started, press Ctrl+Z to exit");
        }
    }

}
