using System;
using System.Collections;
using System.Diagnostics;
using System.Data.OleDb;
using System.Globalization;
using System.Timers;
using System.ComponentModel;
using PLCTools.Service;
using System.Collections.Generic;

namespace PLCTools.Components
{
    public class DspBatch
    {
        /// <summary>
        /// Required designer variable.
        ///
        /// </summary>
        ///
        #region Declare class members

        private Container components = null;
        private OleDbConnection dboleSourceConnection;
        private OleDbConnection dboleTargetConnection;
        private Timer Timer1;
        private BaseClass baseService = new BaseClass();
        private int ii_process_no { get; set; } = 0;
        private DateTime _plcbeat_dateprev;
        private DateTime _plcbeat_datecurr;
        private string _noact_email_sent;
        private DateTime _plccntwrdateprev = DateTime.Parse("1990-01-01");
        private int _plccntwrprev = 0;
        private string _plcbeat_mailsent = "N";
        private string _iesobeat_mailsent = "N";
        private string _pwronbeat_mailsent = "N";
        private string _20tomintue_remind = "N";
        private string _cntloff_email_sent = "N";
        private Int32 _ieso_unix_time;
        private DateTime _ieso_unix_date_time;
        private string _ormarket = "";
        private int _plccountup = 0;
        private int _plccountdown = 0;
        private DateTime _plccntupdateprev = DateTime.Parse("1990-01-01");
        private DateTime _plccntdowndateprev = DateTime.Parse("1990-01-01");
        private string _flink_cntup_mailsent = "N";
        private string _flink_cntdown_mailsent = "N";
        private int _dspcntdownminprev;
        private int _dspcntdownsecprev;
        private int _dspestmincount;
        private int _dspfromcount;

        #endregion

        #region declare member properties

        public DateTime PlcCntWrDatePrev
        {
            get { return _plccntwrdateprev; }
            set { _plccntwrdateprev = value; }
        }

        public int PlcCntWrPrev
        {
            get { return _plccntwrprev; }
            set { _plccntwrprev = value; }
        }

        public string PlcBeatMailSent
        {
            get { return _plcbeat_mailsent; }
            set { _plcbeat_mailsent = value; }
        }

        public string IesoBeatMailSent
        {
            get { return _iesobeat_mailsent; }
            set { _iesobeat_mailsent = value; }
        }

        public string Minute20Remind
        {
            get { return _20tomintue_remind; }
            set { _20tomintue_remind = value; }
        }

        public string CntlOffMailSent
        {
            get { return _cntloff_email_sent; }
            set { _cntloff_email_sent = value; }
        }

        public Int32 IESOUnixTime
        {
            get { return _ieso_unix_time; }
            set { _ieso_unix_time = value; }
        }

        public DateTime IESOUnixDateTime
        {
            get { return _ieso_unix_date_time; }
            set { _ieso_unix_date_time = value; }
        }

        public string IesoOrMarket
        {
            get { return _ormarket; }
            set { _ormarket = value; }
        }

        public int PlcCountUp
        {
            get { return _plccountup; }
            set { _plccountup = value; }
        }

        public int PlcCountDown
        {
            get { return _plccountdown; }
            set { _plccountdown = value; }
        }

        public DateTime PlcCntUpDatePrev
        {
            get { return _plccntupdateprev; }
            set { _plccntupdateprev = value; }
        }

        public DateTime PlcCntDownDatePrev
        {
            get { return _plccntdowndateprev; }
            set { _plccntdowndateprev = value; }
        }

        public string FlinkCntUpMailSent
        {
            get { return _flink_cntup_mailsent; }
            set { _flink_cntup_mailsent = value; }
        }

        public string FlinkCntDownMailSent
        {
            get { return _flink_cntdown_mailsent; }
            set { _flink_cntdown_mailsent = value; }
        }

        public string PwrOnBeatMailSent
        {
            get { return _pwronbeat_mailsent; }
            set { _pwronbeat_mailsent = value; }
        }

        public int DspCountDownMinPrev
        {
            get { return _dspcntdownminprev; }
            set { _dspcntdownminprev = value; }
        }

        public int DspCountDownSecPrev
        {
            get { return _dspcntdownsecprev; }
            set { _dspcntdownsecprev = value; }
        }

        public int DspEstMinCount
        {
            get { return _dspestmincount; }
            set { _dspestmincount = value; }
        }

        public int DispatchFromCount
        {
            get { return _dspfromcount; }
            set { _dspfromcount = value; }
        }

        public OleDbConnection oleSurceConnection
        {
            get
            {
                if (dboleSourceConnection == null)
                {
                    DbCommon dbcomm = new DbCommon();
                    dboleSourceConnection = dbcomm.GetOleConnection("SOURCE");
                    dbcomm = null;
                }

                return dboleSourceConnection;
            }
        }

        public OleDbConnection oleTargetCnnection
        {
            get
            {
                if (dboleTargetConnection == null)
                {
                    DbCommon dbcomm = new DbCommon();
                    dboleTargetConnection = dbcomm.GetOleConnection("TARGET");
                    dbcomm = null;
                }

                return dboleTargetConnection;
            }
        }

        public DateTime prevPLCBeatDate
        {
            get { return _plcbeat_dateprev; }
            set { _plcbeat_dateprev = value; }
        }

        public DateTime currPLCBeatDate
        {
            get { return _plcbeat_datecurr; }
            set { _plcbeat_datecurr = value; }
        }

        public string noActEmailSent
        {
            get { return _noact_email_sent; }
            set { _noact_email_sent = value; }
        }

        #endregion

        public DspBatch()
        {
            // This call is required by the Windows.Forms Component Designer.
            InitializeComponent();
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new Container();
        }


        /// <summary>
        /// Set things in motion so your service can do its work.
        /// </summary>
        public void OnStart()
        {
            // TODO: Add code here to start your service.
            Panel.dspBatchLog.Enqueue("Application Starting...");
            Panel.dspBatchMessage.Enqueue("Application Starting...");
            Panel.dspBatchMail.Enqueue("Application Starting...");
            CreateTimer();
            this.PlcCntWrDatePrev = baseService.getStndSummerDateTime();
            this.PlcCntDownDatePrev = baseService.getStndSummerDateTime();
            this.PlcCntUpDatePrev = baseService.getStndSummerDateTime();
        }

        /// <summary>
        /// Stop this service.
        /// </summary>
        public void OnStop()
        {
            // TODO: Add code here to perform any tear-down necessary to stop your service.
            if (Timer1 != null) Timer1.Enabled = false;
            baseService.actionLog("DspBatch.cs:OnStop", "timer event has been stoped.");
            Panel.dspBatchMail.Enqueue("DspBatch has been stoped.");
            Panel.dspBatchMessage.Enqueue("DspBatch has been stoped.");
        }

        private void CreateTimer()
        {
            int li_second_setting = 0;

            Timer1 = new Timer();
            Timer1.Enabled = true;

            li_second_setting = Int32.Parse(Panel.strTimerSecond);

            Timer1.Interval = (1000) * (li_second_setting);
            Timer1.Elapsed += new ElapsedEventHandler(Timer1_Elapsed);
            Timer1.Start();
        }

        private void Timer1_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                if (ii_process_no == 0)
                {
                    dspBatchProcess("Timer", "from Timer event");
                }
                else
                {
                    baseService.actionLog("DspBatch.cs:Timer1_Elapsed:", "Nothing to do in timer event - value: " + ii_process_no.ToString());
                }
            }

            catch (Exception ex)
            {
                ii_process_no = 0;
                StackTrace st = new StackTrace(new StackFrame(true));
                baseService.actionLog("DspBatch.cs.Timer1_Elapsed(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "timer event exception: " + ex.ToString());
            }
        }

        /// <summary>
        /// Dispatch Load Batch Process
        /// </summary>
        /// <param name="source"></param>
        /// <param name="event_desc"></param>
        ///
        public void dspBatchProcess(string source, string event_desc)
        {
            ClientMsg clientmsg = new ClientMsg();
            string ls_strinf_datetime = "";

            DateTime ld_datetime = DateTime.Parse("Jan01,1990");
            ArrayList parmList = new ArrayList();
            ArrayList resultList = new ArrayList();

            int li_return = 0;
            string ls_emailSubject = "";
            string ls_emailBody = "";
            string ls_msg = "";
            string ls_win_msg_send_ind = "";

            parmList.Add(this.PlcCntWrDatePrev);
            parmList.Add(this.PlcCntWrPrev);
            parmList.Add(this.PlcBeatMailSent);
            parmList.Add(this.IesoBeatMailSent);
            parmList.Add(this.Minute20Remind);
            parmList.Add(this.CntlOffMailSent);
            parmList.Add(this.IESOUnixTime);
            parmList.Add(this.IESOUnixDateTime);
            parmList.Add(this.IesoOrMarket);
            parmList.Add(this.PlcCountUp);
            parmList.Add(this.PlcCountDown);
            parmList.Add(this.PlcCntUpDatePrev);
            parmList.Add(this.PlcCntDownDatePrev);
            parmList.Add(this.FlinkCntUpMailSent);
            parmList.Add(this.FlinkCntDownMailSent);
            parmList.Add(this.PwrOnBeatMailSent);
            parmList.Add(this.DspCountDownMinPrev);
            parmList.Add(this.DspCountDownSecPrev);
            parmList.Add(this.DspEstMinCount);
            parmList.Add(this.DispatchFromCount);

            if (ii_process_no == 0)
            {
                ii_process_no = 1;
                ld_datetime = baseService.getStndSummerDateTime();

                IFormatProvider format = new CultureInfo("fr-FR", true);
                ls_strinf_datetime = ld_datetime.ToString(format);

                DspProcess dspProcess = new DspProcess(parmList);
                //call main process functions
                try
                {
                    li_return = dspProcess.ProcessMain();
                    baseService.actionLog("DspBatch.cs:dspBatchProcess:", "Call dspProcess.ProcessMain() and return value: " + li_return.ToString());

                    if (li_return >= 0)
                    {
                        baseService.actionLog("DspBatch.cs:dspBatchProcess:", "Call dspProcess.ParmPassBack() and result cont: " + resultList.Count.ToString());
                        resultList = dspProcess.ParmPassBack();

                        if (resultList.Count > 0)
                        {
                            this.PlcCntWrDatePrev = (DateTime)resultList[0];
                            this.PlcCntWrPrev = (int)resultList[1];
                            this.PlcBeatMailSent = (string)resultList[2];
                            this.IesoBeatMailSent = (string)resultList[3];
                            this.Minute20Remind = (string)resultList[4];
                            this.CntlOffMailSent = (string)resultList[5];
                            this.IesoOrMarket = (string)resultList[8];
                            this.PlcCountUp = (int)resultList[9];
                            this.PlcCountDown = (int)resultList[10];
                            this.PlcCntUpDatePrev = (DateTime)resultList[11];
                            this.PlcCntDownDatePrev = (DateTime)resultList[12];
                            this.FlinkCntUpMailSent = (string)resultList[13];
                            this.FlinkCntDownMailSent = (string)resultList[14];
                            this.PwrOnBeatMailSent = (string)resultList[15];
                            this.DspCountDownMinPrev = (int)resultList[16];
                            this.DspCountDownSecPrev = (int)resultList[17];
                            this.DspEstMinCount = (int)resultList[18];
                            this.DispatchFromCount = (int)resultList[19];
                        }

                        //get wind message send indicator
                        ls_win_msg_send_ind = Panel.strWinMsgSendInd.ToUpper();
                        Panel.dspBatchMessage.Enqueue(ls_strinf_datetime);
                        if (ls_win_msg_send_ind == "Y")
                        {
                            clientmsg.SendMsgToServer(ls_strinf_datetime);
                        }

                        if (clientmsg != null)
                        {
                            clientmsg = null;
                        }

                        baseService.actionLog("DspBatch.cs:dspBatchProcess:", source + " " + event_desc);
                    }
                    else
                    {
                        ii_process_no = 0;

                        //send email to default people to tell them somthing is wrong
                        StackTrace st = new StackTrace(new StackFrame(true));
                        baseService.actionLog("DspBatch.cs:dspBatchProcess(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):",
                                                    "Error Calling dspProcess.ProcessMain (), result cont: " + resultList.Count.ToString());

                        ls_msg = "Error Calling dspProcess.ProcessMain () at: " + baseService.getStndSummerDateTime().ToString();
                        ls_emailSubject = "System cannot call dspProcess.ProcessMain ()";
                        ls_emailBody = ls_msg;

                        baseService.sendEmail(ls_emailSubject, ls_emailBody);
                    }
                }

                catch (Exception ex)
                {
                    ii_process_no = 0;
                    StackTrace st = new StackTrace(new StackFrame(true));
                    baseService.actionLog("DspBatch.cs:dspBatchProcess(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Error: " + ex.ToString());
                }

                finally
                {
                    if (dspProcess != null)
                    {
                        dspProcess = null;
                        baseService.actionLog("DspBatch.cs:dspBatchProcess:", "Clean up DspProcess in dspBatchProcess of DspBatch.");
                    }
                    if (clientmsg != null)
                    {
                        clientmsg = null;
                    }
                }

                //temperary set to 1 to only run one time
                //finished the transactions and wait for next transaction comming in.
                ii_process_no = 0;
            }
            else
            {
                if (ii_process_no > 0)
                {
                    return;
                }
            }
        }
        ///<summary>
        ///
        ///</summary>
        ///
    }
}
