using System.Collections.Generic;
using System.Net;
using System.Net.Mail;

namespace ImportScheduleData.Utilities
{
    public static class EmailSender
    {
        public static void SendEmail(string toAddresses, string subject, string mailBody, bool isHTML = true, string fromAddress = "", string toCc = "")
        {
            using (var mm = new MailMessage())
            {
                if (!string.IsNullOrEmpty(toAddresses))
                {
                    string[] emailto = toAddresses.Split(';');
                    foreach (string mto in emailto)
                    {
                        mm.To.Add(mto);
                    }
                }
                if (!string.IsNullOrEmpty(toCc))
                {
                    string[] emailCc = toCc.Split(';');
                    foreach (string mCc in emailCc)
                    {
                        mm.CC.Add(mCc);
                    }
                }
                mm.Subject = subject;
                mm.Body = mailBody;
                mm.IsBodyHtml = isHTML;
                var smtp = new SmtpClient();
                smtp.Host = "smtp.intel.com";
                smtp.Port = 25;
                smtp.EnableSsl = false;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential();
                mm.From = new MailAddress("monitor@cware.ie");
                smtp.Send(mm);
            }
        }
        public static void SendEmailWithAttachments(string toAddresses, string subject, string mailBody, List<string> attachmentFiles, bool isHTML = true,
            string fromAddress = "", string toCc = "")
        {
            using (var mm = new MailMessage())
            {
                if (!string.IsNullOrEmpty(toAddresses))
                {
                    string[] emailto = toAddresses.Split(';');
                    foreach (string mto in emailto)
                    {
                        mm.To.Add(mto);
                    }
                }
                if (!string.IsNullOrEmpty(toCc))
                {
                    string[] emailCc = toCc.Split(';');
                    foreach (string mCc in emailCc)
                    {
                        mm.CC.Add(mCc);
                    }
                }
                foreach (var attachmentFile in attachmentFiles)
                {
                    Attachment attachment = new Attachment(attachmentFile);
                    mm.Attachments.Add(attachment);
                }
                mm.Subject = subject;
                mm.Body = mailBody;
                mm.IsBodyHtml = isHTML;
                var smtp = new SmtpClient();
                smtp.Host = "smtp.intel.com";
                smtp.Port = 25;
                smtp.EnableSsl = false;
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new NetworkCredential();
                mm.From = new MailAddress("monitor@cware.ie");
                smtp.Send(mm);
            }
        }
        // See example of how to use a different account smtpClient
        // var smtp = new SmtpClient();
        // smtp.Host = "smtp.gmail.com";
        // smtp.EnableSsl = true;
        // var NetworkCred = new NetworkCredential(fromEmail, password);
        // smtp.UseDefaultCredentials = false;
        // smtp.Credentials = NetworkCred;
        // smtp.Port = 587;
        public static void SendEmailFromAccount(SmtpClient smtpClient, string toAddresses, string subject, string mailBody, bool isHTML = true)
        {
            using (var mm = new MailMessage())
            {
                mm.To.Add(toAddresses);
                mm.Subject = subject;
                mm.Body = mailBody;
                mm.IsBodyHtml = isHTML;
                smtpClient.Send(mm);
            }
        }
    }
}
