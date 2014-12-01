using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GSS_Alert_Service
{
    public static class Mailer
    {
        private static List<MailMessage> MessageQueue = new List<MailMessage>();

        /// <summary>
        /// Adds a message to the mail queue for transmission
        /// </summary>
        /// <param name="To">To Address</param>
        /// <param name="From">From Address</param>
        /// <param name="Subject">Message Subject</param>
        /// <param name="MessageText">Message text with tokens</param>
        /// <param name="TokenValues">Tokens and definitions</param>
        public static void QueueMessage(string To, string From, string Subject, string MessageText, Dictionary<string, string> Tokens)
        {
            foreach (string key in Tokens.Keys)
            {
                MessageText = MessageText.Replace(key, Tokens[key]);
            }

            MailMessage message = new MailMessage(From, To, Subject, MessageText);

            MessageQueue.Add(message);
        }

        /// <summary>
        /// Adds a message to the mail queue for transmission
        /// </summary>
        /// <param name="To">To Address</param>
        /// <param name="From">From Address</param>
        /// <param name="Subject">Message Subject</param>
        /// <param name="MessageText">Message text with tokens</param>
        /// <param name="TokenValues">Tokens and definitions</param>
        /// /// <param name="CC">CC Address(es)</param>
        public static void QueueMessage(string To, string From, string Subject, string MessageText, Dictionary<string, string> Tokens, string[] CC)
        {
            foreach (string key in Tokens.Keys)
            {
                MessageText = MessageText.Replace(key, Tokens[key]);
            }

            MailMessage message = new MailMessage(From, To, Subject, MessageText);
            if (CC.Length > 0)
            {
                foreach (string address in CC)
                {
                    message.CC.Add(address);
                }
            }

            MessageQueue.Add(message);
        }

        public static int SendMessageQueue()
        {
            int count = MessageQueue.Count;
            Task sendmail = new Task(delegate { send(); });
            sendmail.Start();
            return count;
        }

        private static void send()
        {
            using (SmtpClient mail = new SmtpClient("smtp.vmware.com"))
            {
                foreach (MailMessage Message in MessageQueue)
                {
                    mail.Send(Message);
                }
            }

            MessageQueue.Clear();

        }
    }
}
