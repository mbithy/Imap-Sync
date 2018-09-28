using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AE.Net.Mail;
using AE.Net.Mail.Imap;

namespace ImapSync
{
    class Program
    {
        static void Main(string[] args)
        {
            var vrt = "sercret hehe";
            var split = vrt.Split(':');
            var source= new ImapClient(split[2],split[0],split[1]);
            var destination= new ImapClient(split[6], split[4], split[5], AuthMethods.Login, int.Parse(split[7]), true, true);
            Console.WriteLine("Begining");
            var sourceFolder = source.ListMailboxes(string.Empty, "*");
            var destinationFolder = destination.ListMailboxes(string.Empty, "*");
            foreach (var mailbox in sourceFolder)
            {
                if (!mailbox.Name.Contains("Gmail"))
                {
                    if (FolderExists(mailbox.Name, destinationFolder))
                    {
                        BeginCopy(source, mailbox, destination);
                    }
                    else
                    {
                        destination.CreateMailbox(mailbox.Name);
                        destinationFolder = destination.ListMailboxes(string.Empty, "*");
                        BeginCopy(source, mailbox, destination);
                    }
                }
            }
            Console.WriteLine("Sending Done");
            Console.ReadKey();
        }

        private static void BeginCopy(ImapClient source, Mailbox mailbox, ImapClient destination)
        {
            var srcMsgCount = source.GetMessageCount(mailbox.Name);
            if (srcMsgCount > 0)
            {
                source.SelectMailbox(mailbox.Name);
                Console.WriteLine("Getting source mesages "+srcMsgCount);
                var srcMsgs = source.GetMessages(0, 5,false);
                var destMsgCount = destination.GetMessageCount(mailbox.Name);
                if (destMsgCount > 0)
                {
                    destination.SelectMailbox(mailbox.Name);
                    var destMsgs = destination.GetMessages(0, destMsgCount);
                    CopyEmails(srcMsgs, destMsgs, destination, mailbox);
                }
                else
                {
                    var destMsgs = new AE.Net.Mail.MailMessage[0];
                    CopyEmails(srcMsgs, destMsgs, destination, mailbox);
                }
            }
        }

        private static void CopyEmails(MailMessage[] srcMsgs, MailMessage[] destMsgs, ImapClient destination, Mailbox mailbox)
        {
            foreach (var mailMessage in srcMsgs)
            {
                if (!MessageExists(mailMessage.MessageID, destMsgs))
                {
                    destination.AppendMail(mailMessage, mailbox.Name);
                    Console.WriteLine("Copied " + mailMessage.MessageID);
                }
            }
        }

        static bool FolderExists(string name, Mailbox[] boxes)
        {
            foreach (var mailbox in boxes)
            {
                if (mailbox.Name == name)
                {
                    return true;
                }
            }
            return false;
        }

        static bool MessageExists(string msgId, MailMessage[] messages)
        {
            foreach (var mailMessage in messages)
            {
                if (mailMessage.Uid == msgId)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
