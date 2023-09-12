using Microsoft.Exchange.WebServices.Data;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace SendExchangeMail
{
    class Program
    {
        /*
         This small program sends E-Mails using Microsoft Exchange ActiveSync on Exchange Servers 2013 and newer
        Arguments:
        set-credentials username password corresponding-email@domain.com
        and
        recipient@domain.com;seperatedByColon@domain.com Subject Body /path/to/attachments.docx;/path/to/other/attachment.png

        This Program uses the provided email to find the Exchange server responsible using Autodiscover

        WARNING: The way this tool stores credentials IS NOT SECURE
         
         */
        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Argument error");
                return;
            }

            if (args.Length == 4 && args[0].Equals("set-credentials"))
            {
                ProtectAndStoreCredentials(args[1], args[2], args[3]);
                Console.WriteLine("Stored new credentials");
                return;
            }
            string[] cred = RetrieveAndDecryptCredentials();
            if (cred[0] == null || cred[1] == null || cred[2] == null)
            {
                Console.WriteLine("No Credentials or invalid credentials stored");
                return;
            }

            string user = cred[0];
            string passw = cred[1];
            string domainEmail = cred[2];

            string[] recipient = args[0].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            string subject = args[1].Replace("\"", "").Replace("'", "");
            subject = subject.Length == 0 ? "Empty subject" : subject;

            string body = args[2].Replace("\"", "").Replace("'", "");
            body = body.Length == 0 ? "Empty body" : body;

            string[] attachments = new string[0];
            if (args.Length > 3)
            {
                attachments = args[3].Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            }


            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
            service.Credentials = new WebCredentials(user, passw);

            try
            {
                service.AutodiscoverUrl(domainEmail, RedirectionUrlValidationCallback);
            }
            catch (Exception er) {
                Console.WriteLine("Autodiscover failed. Email: " + domainEmail + " user: ");
                Console.WriteLine(er.Message);
                Console.WriteLine(er.StackTrace);
                return;
            }

            EmailMessage email = new EmailMessage(service);
            email.Subject = subject;
            email.Body = new MessageBody(body);
            foreach (string s in attachments) 
            {
                email.Attachments.AddFileAttachment(s);
            }

            foreach (string s in recipient)
            {
                email.ToRecipients.Add(s);
            }
           
            
            try
            {
                email.SendAndSaveCopy();
            }
            catch (Exception ef)
            {
                Console.WriteLine("Could not send email" );
                Console.WriteLine(ef.Message);
            }
          

        }

     
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            return true;
        }


        public static void ProtectAndStoreCredentials(string username, string password, string email)
        {
           

            // Store encryptedData in the Registry
            Registry.SetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\ExchangeMail", "Password", Convert.ToBase64String(Encoding.UTF8.GetBytes(password)), RegistryValueKind.String);
            Registry.SetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\ExchangeMail", "Username", Convert.ToBase64String(Encoding.UTF8.GetBytes(username)), RegistryValueKind.String);
            Registry.SetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\ExchangeMail", "Email", Convert.ToBase64String(Encoding.UTF8.GetBytes(email)), RegistryValueKind.String);

        }

        public static string[] RetrieveAndDecryptCredentials()
        {
            string[] ret = new string[3];

            string passwordData = Encoding.UTF8.GetString(Convert.FromBase64String((string)Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\ExchangeMail", "Password", null)));


            string usernameData = Encoding.UTF8.GetString(Convert.FromBase64String((string)Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\ExchangeMail", "Username", null)));


            string emailData = Encoding.UTF8.GetString(Convert.FromBase64String((string)Registry.GetValue("HKEY_LOCAL_MACHINE\\SOFTWARE\\ExchangeMail", "Email", null)));

            ret[0] = usernameData;
            ret[1] = passwordData;
            ret[2] = emailData;
          

            return ret;
        }

    }



}
