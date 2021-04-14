using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using Exception = System.Exception;

namespace Carbuncle.Commands
{
    public class MailSender
    {
        public MailSender()
        {

        }
        public void SendEmail(string[] recipients, string body, string subject)
        {
            Console.WriteLine("[+] Sending an e-mail.\r\nRecipients: {0}\r\nSubject: {1}\r\nBody: {2}", String.Join(",", recipients), subject, body);
            try
            {
                Application outlookApplication = new Application();
                MailItem msg = (MailItem)outlookApplication.CreateItem(OlItemType.olMailItem);
                msg.HTMLBody = body;
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
        public void SendEmail(string[] recipients, string body, string subject, string attachment, string attachmentname)
        {
            Console.WriteLine("[+] Sending an e-mail.\r\nRecipients: {0}\r\nSubject: {1}\r\nAttachment Path: {2}\r\nAttachment Name: {3}\r\nBody: {4}", String.Join(",", recipients), subject, attachment, attachmentname, body);
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
    }
}
