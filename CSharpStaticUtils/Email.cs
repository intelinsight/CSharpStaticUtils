using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;

namespace CSharpStaticUtils
{
    public class Email
    {
        public static bool SendEmail(string[] to, string[] cc, string from, string subject, string body, string attachmentFilePath, string login, string password)
        {
            bool success = false;
            try
            {
                Console.WriteLine(string.Format("Sending email to {0}" + Environment.NewLine, to));
                MailMessage email = new System.Net.Mail.MailMessage();
                foreach (string id in to)
                {
                    email.To.Add(id);
                }
                foreach (string id in cc)
                {
                    email.CC.Add(id);
                }
                email.Subject = subject;
                email.From = new System.Net.Mail.MailAddress(from);
                email.Body = body;
                if (!string.IsNullOrEmpty(attachmentFilePath))
                {
                    Attachment newsletter = new Attachment(attachmentFilePath);
                    email.Attachments.Add(newsletter);
                }
                email.IsBodyHtml = false;
                System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.office365.com");
                smtp.EnableSsl = true;
                smtp.Credentials = new System.Net.NetworkCredential(login, password);
                smtp.Port = 587;
                smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtp.Send(email);
                success = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("Email Error; To={0}; Error={1}", to, ex.ToString()));
            }
            return success;
        }
    }
}
