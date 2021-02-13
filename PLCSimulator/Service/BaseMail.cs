using System;
using System.Diagnostics;
using System.Net.Mail;

namespace PLCTools.Service
{
    /// <summary>

    public class BaseMail : BaseClass
    {
        private MailMessage msgMail;
        private SmtpClient smtpClient;

        public BaseMail()
        {
            //
            // TODO: Add constructor logic here
            //
            smtpClient = new SmtpClient(Panel.strMailHost, Panel.strMailPort);
            smtpClient.EnableSsl = Panel.isSSLMail;
        }


        ~BaseMail()
        {
            if (msgMail != null)
            {
                msgMail = null;
            }
        }

        public int sendEmail(System.Collections.ArrayList mailToList, string mailsubject, string mailContents)
        {
            int li_sentmail = 0;
            string toMailAddr = "";

            DateTime ldt_datetime;
            ldt_datetime = getStndSummerDateTime();

            if (mailContents != null)
            {
                mailContents = mailContents + "\nsent at: " + ldt_datetime.ToString();
            }
            if (msgMail != null)
            {
                try
                {
                    if (mailToList.Count > 0)
                    {
                        msgMail = new MailMessage();
                        foreach (string address in mailToList)
                        {
                            msgMail.From = new MailAddress(Panel.strMailFrom);
                            msgMail.To.Add(address);
                        }
                    }
                    else
                    {
                        msgMail = new MailMessage(Panel.strMailFrom, Panel.strMailTo);
                    }
                    msgMail = new MailMessage(Panel.strMailFrom, toMailAddr);
                    msgMail.Subject = mailsubject;
                    msgMail.Body = mailContents;
                    Panel.dspBatchMail.Enqueue(ldt_datetime.ToString("MMM.dd") + " " + ldt_datetime.ToShortTimeString() + "\t" + mailsubject);
                    if (Panel.isSendMail)
                    {
                        if (Panel.strMailPass != "") smtpClient.Credentials = new System.Net.NetworkCredential(Panel.strMailLogin, Panel.strMailPass); 
                        smtpClient.Send(msgMail);
                    }
                    actionLog("BaseMail:sendEmail:", "Email sent to: " + toMailAddr);

                    li_sentmail = 1;
                }

                catch (Exception ex)
                {
                    StackTrace st = new StackTrace(new StackFrame(true));
                    actionLog("BaseMail:sendEmail (" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Email error: " + ex.ToString());

                    li_sentmail = -1;
                }
            }
            else
            {
                toMailAddr = "";
                try
                {
                    if (mailToList.Count > 0)
                    {
                        msgMail = new MailMessage();
                        foreach (string address in mailToList)
                        {
                            msgMail.From = new MailAddress(Panel.strMailFrom);
                            msgMail.To.Add(address);
                        }
                    }
                    else
                    {
                        msgMail = new MailMessage(Panel.strMailFrom, Panel.strMailTo);
                    }
                    msgMail.Subject = mailsubject;
                    msgMail.Body = mailContents + "\nsent at: " + ldt_datetime.ToString();
                    Panel.dspBatchMail.Enqueue(ldt_datetime.ToString("MMM.dd") + " " + ldt_datetime.ToShortTimeString() + "\t" + mailsubject);
                    if (Panel.isSendMail)
                    {
                        if (Panel.strMailPass != "") smtpClient.Credentials = new System.Net.NetworkCredential(Panel.strMailLogin, Panel.strMailPass);
                        smtpClient.Send(msgMail);
                    }
                    actionLog("BaseMail:sendEmail:", "Email sent to: " + toMailAddr);

                    li_sentmail = 1;
                }

                catch (Exception ex)
                {
                    StackTrace st = new StackTrace(new StackFrame(true));
                    actionLog("BaseMail:sendEmail(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Email error: " + ex.ToString());

                    li_sentmail = -1;
                }
            }

            return li_sentmail;
        }

        public void sendDefaultEmail(string subject, string mailMsg)
        {
            DateTime ldt_datetime;

            ldt_datetime = getStndSummerDateTime();

            if (msgMail == null)
            {
                msgMail = new MailMessage();
                msgMail.Priority = MailPriority.Normal;
            }

            try
            {
                msgMail = new MailMessage(Panel.strMailFrom, Panel.strMailTo);
                msgMail.Subject = subject;
                msgMail.Body = mailMsg + "\nsent at: " + ldt_datetime.ToString();
                Panel.dspBatchMail.Enqueue(ldt_datetime.ToString("MMM.dd") + " " + ldt_datetime.ToShortTimeString() + "\t" + subject);
                if (Panel.isSendMail)
                {
                    if (Panel.strMailPass != "") smtpClient.Credentials = new System.Net.NetworkCredential(Panel.strMailLogin, Panel.strMailPass);
                    smtpClient.Send(msgMail);
                }
                actionLog("BaseMail:sendDefaultEmail:", "Default email sent to :" + msgMail.To + " with Message: " + mailMsg);
            }

            catch (Exception ex)
            {
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("BaseMail:sendDefaultEmail(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Email error: " + ex.ToString());
            }
        }

    }
}
