using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net;

namespace AuthorizeTransactions
{
    class Email
    {

        static public bool SendEmail(string file = "") {

            try {
                
                string emailList = Properties.Settings.Default.emailList;
                //string[] disList = new string[4] { "enio.maiale@gmail.com", "heda@mibrk.com", "aotatti@mibrk.com", "dalia@mibrk.com" };
                char[] delimiter1 = new char[] { ',' };
                string[] disList = emailList.Split(delimiter1, StringSplitOptions.RemoveEmptyEntries);


                SmtpClient client = new SmtpClient();
                client.Port = 587;
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Host = "smtp.gmail.com";
                client.Credentials = new NetworkCredential("", "");

                foreach (string email in disList) {
                    try
                    {
                        MailMessage mail = new MailMessage("enio.maiale@gmail.com", email);
                        mail.Subject = "Daily Credit Card Report - Escrow";
                        mail.Body = "Please open the excel file attached.";
                        mail.Attachments.Add(new Attachment(file));
                        client.Send(mail);
                    }
                    catch (Exception es) { 
                        
                    }
                }

                return true;
            }
            catch (Exception es) {
                return false;
            }

        }
    }
}
