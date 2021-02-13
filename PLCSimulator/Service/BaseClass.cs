using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Data.OleDb;
using System.Windows.Forms;

namespace PLCTools.Service
{
    /// <summary>
    /// Summary description for BaseClass.
    /// </summary>
    public class BaseClass
    {
        public BaseClass()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        public int sendDbEmail(string dspType, string dspScenario, string emailsubject, string emailBody)
        {
            ArrayList parmList = new ArrayList();
            ArrayList parmTypeList = new ArrayList();
            ArrayList parmValueList = new ArrayList();
            ArrayList parmSizeList = new ArrayList();
            parmList.Add("@dsp_type");
            parmList.Add("@scenario");
            parmValueList.Add(dspType);
            parmValueList.Add(dspScenario);
            parmTypeList.Add(OleDbType.Char);
            parmTypeList.Add(OleDbType.Char);
            parmSizeList.Add(16);
            parmSizeList.Add(16);
            int li_return = 1;
            int li_emailsent = 0;
            try
            {
                li_emailsent = sendEmail("ssp_get_dsp_email_addr", parmList, parmTypeList, parmSizeList, parmValueList, emailsubject, emailBody);
                if (li_emailsent >= 0)
                {
                    actionLog("BaseClass:sendDbEmail:", "The following has been sent via email: " + emailBody);
                }
                else
                {
                    li_emailsent = sendEmail(emailsubject, emailBody);
                    actionLog("BaseClass:sendDbEmail:", "The following has been sent via email (default address): " + emailBody);
                }
            }
            catch (Exception ex)
            {
                li_return = -1;
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("BaseClass:sendDbEmail(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Unable to Send Email: " + ex.ToString());
            }
            return li_return;
        }
        public int sendEmail(string proName, ArrayList parmName, ArrayList parmType, ArrayList parmSize, ArrayList parmValue, string subject, string msgText)
        {
            ArrayList emailList = new ArrayList();
            ArrayList parmList = new ArrayList();
            DataTable dt;
            DataTable dt1;
            dt = new DataTable();
            dt1 = new DataTable();
            DbCommon dbcomm = new DbCommon();
            BaseMail msgMail = new BaseMail();
            string ls_actscenario = "";
            string ls_act_count = "";
            int li_return = 0;
            int li_no_email_timefrom = 0;
            int li_no_email_timeto = 0;
            int li_curr_hour = 0;
            int li_num_extend = 0;
            int li_hour_limit_value = 0;
            //1.get email address
            try
            {
                this.actionLog("BaseClass:sendEmail:", "Send Email Procedure Name: " + proName);
                OleDbConnection cnnSql = dbcomm.GetOleConnection("TARGET");
                parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                dt = dbcomm.getOleProData(proName, parmList, cnnSql);
                dt1 = dbcomm.getOleTableData("select dsp_type,dsp_sub_type,dsp_act_desc,action_count,hour_count from Dsp_Est_Action_Count", cnnSql);
                if (cnnSql != null)
                {
                    cnnSql.Dispose();
                }
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        this.actionLog("BaseClass:sendEmail:", "get no email hours: " + myRow.ItemArray[3].ToString() + ";" + myRow.ItemArray[4].ToString() + ";" + parmValue[0].ToString());
                        //get numbr of extention and hour limit value
                        li_num_extend = Int32.Parse(myRow.ItemArray[5].ToString());
                        li_hour_limit_value = Int32.Parse(myRow.ItemArray[6].ToString());
                        this.actionLog("BaseClass:sendEmail:", "get No. Extend: " + myRow.ItemArray[5].ToString() + ";" + myRow.ItemArray[6].ToString());
                        //check no email hours
                        li_curr_hour = DateTime.Now.Hour;
                        if (li_curr_hour == 0)
                        {
                            li_curr_hour = 24;
                        }
                        if (myRow.ItemArray[3].ToString() != "")
                        {
                            li_no_email_timefrom = Int32.Parse(myRow.ItemArray[3].ToString().Substring(0, 2));
                            if (li_no_email_timefrom == 0)
                            {
                                li_no_email_timefrom = 24;
                            }
                            if (myRow.ItemArray[4].ToString() != "")
                            {
                                li_no_email_timeto = Int32.Parse(myRow.ItemArray[4].ToString().Substring(0, 2));
                                if (li_no_email_timeto == 0)
                                {
                                    li_no_email_timeto = 24;
                                }
                                //----------------------------------------------------------
                                // check curr hour and no email time range to decide whether this 
                                // email should added to list or not.
                                //NO EMAIL TIME RANGE ACROSS THE DAY!!!
                                if (li_no_email_timeto <= li_no_email_timefrom)
                                {
                                    //the curr time is out of no email time range, send email.
                                    if ((li_curr_hour < li_no_email_timefrom) & (li_curr_hour > li_no_email_timeto))
                                    {
                                        emailList.Add(myRow.ItemArray[2].ToString());
                                    }
                                    else
                                    {
                                        this.actionLog("BaseClass:sendEmail:", "Check DISPATCHON in no email time range across day");
                                        //NO EMAIL NEED TO BE SENT EXCEPT DISPATCHON AND TIMELIMIT
                                        if ((parmValue[0].ToString().ToUpper() == "DISPATCHON") & (parmValue[1].ToString().ToUpper() == "TIMELIMIT"))
                                        {
                                            if (dt1.Rows.Count > 0)
                                            {
                                                foreach (DataRow myRow1 in dt1.Rows)
                                                {
                                                    ls_actscenario = myRow1.ItemArray[0].ToString();
                                                    ls_act_count = myRow1.ItemArray[3].ToString();
                                                    if (((myRow1.ItemArray[0].ToString()).ToUpper() == "BIDCHANGE") & ((myRow1.ItemArray[1].ToString()).ToUpper() == "BIDCHANGE"))
                                                    {
                                                        if ((Int32.Parse(myRow1.ItemArray[3].ToString()) > li_num_extend) ||
                                                            (Int32.Parse(myRow1.ItemArray[4].ToString()) > li_hour_limit_value))
                                                        {
                                                            emailList.Add(myRow.ItemArray[2].ToString());
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    //NO EMAIL TIME RNGE IN THE SAME DAY.
                                    //curr time is out of no email time range, send email.
                                    if ((li_curr_hour < li_no_email_timefrom) || (li_curr_hour > li_no_email_timeto))
                                    {
                                        //send email
                                        emailList.Add(myRow.ItemArray[2].ToString());
                                    }
                                    else
                                    {
                                        this.actionLog("BaseClass:sendEmail:", "Check DISPATCHON in no email time range in the same day");
                                        //there is should no email exception for "DISPATCHON AND TIMELIMIT"
                                        //check "DISPATCHON AND TIMELIMIT"
                                        if ((parmValue[0].ToString().ToUpper() == "DISPATCHON") & (parmValue[1].ToString().ToUpper() == "TIMELIMIT"))
                                        {
                                            if (dt1.Rows.Count > 0)
                                            {
                                                foreach (DataRow myRow1 in dt1.Rows)
                                                {
                                                    ls_actscenario = myRow1.ItemArray[0].ToString();
                                                    ls_act_count = myRow1.ItemArray[3].ToString();
                                                    if (((myRow1.ItemArray[0].ToString()).ToUpper() == "BIDCHANGE") & ((myRow1.ItemArray[1].ToString()).ToUpper() == "BIDCHANGE"))
                                                    {
                                                        if ((Int32.Parse(myRow1.ItemArray[3].ToString()) > li_num_extend) ||
                                                            (Int32.Parse(myRow1.ItemArray[4].ToString()) > li_hour_limit_value))
                                                        {
                                                            emailList.Add(myRow.ItemArray[2].ToString());
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            //there is no no email time range limit. send email
                            emailList.Add(myRow.ItemArray[2].ToString());
                        }
                        this.actionLog("BaseClass:sendEmail:", "Scenario Value: " + ls_actscenario + ";" + ls_act_count);
                        this.actionLog("BaseClass:sendEmail:", "no email hours: " + li_no_email_timefrom.ToString() + ";" + li_no_email_timeto.ToString());
                        this.actionLog("BaseClass:sendEmail:", "Get email address for type " + subject + ":" + myRow.ItemArray[2].ToString());
                    }
                }
                this.actionLog("BaseClass:sendEmail:", "The email list before sendemail: ");
                for (int i = 0; i < emailList.Count; i++)
                {
                    this.actionLog("BaseClass:sendEmail:", emailList[i].ToString());
                }
                if (emailList.Count > 0)
                {
                    li_return = msgMail.sendEmail(emailList, subject, msgText);
                }
                else
                {
                    this.actionLog("BaseClass:sendEmail:", "No email list, send default email.");
                    li_return = msgMail.sendEmail(subject, msgText);
                }
                this.actionLog("BaseClass:sendEmail:", "Email send out status value returned from msgMail: " + li_return.ToString());
            }
            catch (Exception ex)
            {
                li_return = -1;
                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("BaseClass:sendEmail(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception error: " + ex.ToString());
            }
            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }
                dbcomm = null;
                msgMail = null;
            }
            return li_return;
        }
        public int sendEmail(string subject, string msgText)
        {
            int li_return = 0;
            BaseMail msgMail = new BaseMail();
            try
            {
                msgMail.sendDefaultEmail(subject, msgText);
                li_return = 1;
            }

            catch (Exception ex)
            {
                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("BaseClass:sendEmail(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception error: " + ex.ToString());
                li_return = -1;
            }
            finally
            {
                if (msgMail != null)
                {
                    msgMail = null;
                }
            }
            return li_return;
        }
        public void actionLog(string source, string actPath, string event_desc)
        {
            string actionLogPath = "";
            string actionLogPathOrg = "";
            string actionLogDeletePath = "";
            double logDeleteDays = 7;
            DateTime ldt_deleteDate = DateTime.Parse("Jan01,1990");
            actionLogPathOrg = Panel.strActionLogPath;
            try
            {
                actionLogPath = actionLogPathOrg + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00") + DateTime.Now.Hour.ToString("00") + ".log";
                ldt_deleteDate = DateTime.Now.AddDays((-1) * logDeleteDays);
                for (int i = 0; i < 24; i++)
                {
                    actionLogDeletePath = actionLogPathOrg + ldt_deleteDate.Year.ToString() + ldt_deleteDate.Month.ToString("00") + ldt_deleteDate.Day.ToString("00") + i.ToString("00") + ".log";
                    if (File.Exists(actionLogDeletePath))
                    {
                        File.Delete(actionLogDeletePath);
                    }
                }
                FileStream fileStream = new FileStream(actionLogPath, FileMode.Append, FileAccess.Write, FileShare.Write);
                TextWriter tw = new StreamWriter(fileStream);
                tw.WriteLine(getStndSummerDateTime().ToString() + " " + source + " " + event_desc);
                tw.Close();
            }
            catch (Exception ex)
            {
                string ls_error = ex.ToString();
            }
        }
        public void actionLog(string source, string event_desc)
        {
            string actionLogPath = "";
            string actionLogPathOrg = "";
            string actionLogDeletePath = "";
            double logDeleteDays = 7;
            DateTime ldt_deleteDate = DateTime.Parse("Jan01,1990");
            actionLogPathOrg = Panel.strActionLogPath;
            try
            {
                actionLogPath = actionLogPathOrg + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00") + ".log";
                actionLogPath = actionLogPathOrg + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00") + DateTime.Now.Hour.ToString("00") + ".log";
                ldt_deleteDate = DateTime.Now.AddDays((-1) * logDeleteDays);
                for (int i = 0; i < 24; i++)
                {
                    actionLogDeletePath = actionLogPathOrg + ldt_deleteDate.Year.ToString() + ldt_deleteDate.Month.ToString("00") + ldt_deleteDate.Day.ToString("00") + i.ToString("00") + ".log";
                    if (File.Exists(actionLogDeletePath))
                    {
                        File.Delete(actionLogDeletePath);
                    }
                }
                FileStream fileStream = new FileStream(actionLogPath, FileMode.Append, FileAccess.Write, FileShare.Write);
                TextWriter tw = new StreamWriter(fileStream);
                tw.WriteLine(getStndSummerDateTime().ToString() + " " + source + " " + event_desc);
                Panel.dspBatchLog.Enqueue(source + " " + event_desc);
                tw.Close();
            }
            catch (Exception ex)
            {
                string ls_test = ex.ToString();
            }
        }
        public int getComputerName(ref string ps_computerName, ref string ps_login)
        {
            int li_return = 1;
            try
            {
                ps_computerName = SystemInformation.ComputerName.ToString();
                ps_login = SystemInformation.UserName.ToString();
            }
            catch (Exception ex)
            {
                li_return = -1;
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("BaseClass:getComputerName(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception error: " + ex.ToString());
            }
            return li_return;
        }
        public DateTime getStndSummerDateTime()
        {
            int li_curr_dst_year;
            int li_day;
            DateTime ldt_curr_dst_date_start = DateTime.Parse("1990-01-01");
            DateTime ldt_curr_dst_date_end = DateTime.Parse("1990-01-01");
            DateTime ldt_datetime = DateTime.Parse("1990-01-01");
            //check whether this is summer time.
            //check current year DST start
            li_curr_dst_year = DateTime.Now.Year;
            //current DST start date
            for (int li_cont = 1; li_cont <= 10; li_cont++)
            {
                //ranges from zero, indicating Sunday, to six, indicating Saturday
                li_day = (int)DateTime.Parse(li_curr_dst_year.ToString() + "-" + "04" + "-" + li_cont.ToString()).DayOfWeek;
                if (li_day == 0)
                {
                    ldt_curr_dst_date_start = DateTime.Parse(li_curr_dst_year.ToString() + "-" + "04" + "-" + li_cont.ToString() + " 02:00:00 AM");
                    break;
                }
            }
            //current DST end date
            for (int li_cont = 15; li_cont <= 31; li_cont++)
            {
                //"10" Octobor
                li_day = (int)DateTime.Parse(li_curr_dst_year.ToString() + "-" + "10" + "-" + li_cont.ToString()).DayOfWeek;
                if (li_day == 0)
                {
                    ldt_curr_dst_date_end = DateTime.Parse(li_curr_dst_year.ToString() + "-" + "10" + "-" + li_cont.ToString() + " 02:00:00 AM");
                }
            }
            if ((DateTime.Now >= ldt_curr_dst_date_start) && (DateTime.Now <= ldt_curr_dst_date_end))
            {
                ldt_datetime = DateTime.Now.AddHours(1);
            }
            else
            {
                ldt_datetime = DateTime.Now;
            }
            return ldt_datetime;
        }
    }
}
