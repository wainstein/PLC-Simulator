using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;

namespace PLCTools.Service
{
    /// <summary>
    /// Summary description for DspProcess.
    /// </summary>
    public class DspProcess : BaseClass
    {
        #region declare instance variables

        const string ls_oractive = "ORActivation - Shutdown ASAP";
        const string ls_engoff = "Dispatch Off";
        const string ls_engon = "Dispatch On";

        OleDbConnection oleCnn;
        OleDbConnection oleSrcCnn;
        DbCommon dbcomm;
        String str = new String(' ', 40);

        private DateTime _plccntwrdateprev;
        private int _plccntwrprev;
        private string _plcbeat_mailsent;
        private string _iesobeat_mailsent;
        private string _pwronbeat_mailsent;
        private string _20tomintue_remind;
        private string _cntloff_email_sent;
        private Int32 _ieso_unix_time;
        private DateTime _ieso_unix_date_time;
        private string _ormarket = "";
        private int _plccountup;
        private int _plccountdown;
        private DateTime _plccntupdateprev;
        private DateTime _plccntdowndateprev;
        private string _flink_cntup_mailsent;
        private string _flink_cntdown_mailsent;

        private int _dspcntdownminprev;
        private int _dspcntdownsecprev;
        private int _dspestmincount;
        private int _dspfromcount;

        #endregion

        #region declare propertities
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

        #endregion

        public DspProcess()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        public DspProcess(ArrayList pa_parmList)
        {
            //
            // TODO: Add constructor logic here
            //
            //Assign parameters value to propertites.
            if (pa_parmList.Count > 4)
            {
                this.PlcCntWrDatePrev = (DateTime)pa_parmList[0];
                this.PlcCntWrPrev = (int)pa_parmList[1];
                this.PlcBeatMailSent = (string)pa_parmList[2];
                this.IesoBeatMailSent = (string)pa_parmList[3];
                this.Minute20Remind = (string)pa_parmList[4];
                this.CntlOffMailSent = (string)pa_parmList[5];
                this.IESOUnixTime = (Int32)pa_parmList[6];
                this.IESOUnixDateTime = (DateTime)pa_parmList[7];
                this.IesoOrMarket = (string)pa_parmList[8];
                this.PlcCountUp = (Int32)pa_parmList[9];
                this.PlcCountDown = (Int32)pa_parmList[10];
                this.PlcCntUpDatePrev = (DateTime)pa_parmList[11];
                this.PlcCntDownDatePrev = (DateTime)pa_parmList[12];
                this.FlinkCntUpMailSent = (string)pa_parmList[13];
                this.FlinkCntDownMailSent = (string)pa_parmList[14];
                this.PwrOnBeatMailSent = (string)pa_parmList[15];
                this.DspCountDownMinPrev = (Int32)pa_parmList[16];
                this.DspCountDownSecPrev = (Int32)pa_parmList[17];
                this.DspEstMinCount = (Int32)pa_parmList[18];
                this.DispatchFromCount = (Int32)pa_parmList[19];

                actionLog("DspProcess.cs:DspProcess:", "Parameters for DspProcess------------------------------------");
                actionLog("DspProcess.cs:DspProcess:", "PlcCntWrDatePrev: " + this.PlcCntWrDatePrev.ToString());
                actionLog("DspProcess.cs:DspProcess:", "PlcCntWrPrev: " + this.PlcCntWrPrev.ToString());
                actionLog("DspProcess.cs:DspProcess:", "PlcBeatMailSent: " + this.PlcBeatMailSent.ToString());
                actionLog("DspProcess.cs:DspProcess:", "IesoBeatMailSent: " + this.IesoBeatMailSent.ToString());
                actionLog("DspProcess.cs:DspProcess:", "Minute20Remind: " + this.Minute20Remind.ToString());
                actionLog("DspProcess.cs:DspProcess:", "CntlOffMailSent: " + this.CntlOffMailSent.ToString());
                actionLog("DspProcess.cs:DspProcess:", "IESOUnixTime: " + this.IESOUnixTime.ToString());
                actionLog("DspProcess.cs:DspProcess:", "IESOUnixDateTime: " + this.IESOUnixDateTime.ToString());
                actionLog("DspProcess.cs:DspProcess:", "IesoORMarket: " + this.IesoOrMarket.ToString());
                actionLog("DspProcess.cs:DspProcess:", "PlcCountUp: " + this.PlcCountUp.ToString());
                actionLog("DspProcess.cs:DspProcess:", "PlcCountDown: " + this.PlcCountDown.ToString());
                actionLog("DspProcess.cs:DspProcess:", "PlcCntUpDatePrev: " + this.PlcCntUpDatePrev.ToString());
                actionLog("DspProcess.cs:DspProcess:", "PlcCntDownDatePrev: " + this.PlcCntDownDatePrev.ToString());
                actionLog("DspProcess.cs:DspProcess:", "FlinkCntUpMailSent: " + this.FlinkCntUpMailSent.ToString());
                actionLog("DspProcess.cs:DspProcess:", "FlinkCntUpMailSent: " + this.FlinkCntDownMailSent.ToString());
                actionLog("DspProcess.cs:DspProcess:", "PwrOnBeatMailSent: " + this.PwrOnBeatMailSent.ToString());
                actionLog("DspProcess.cs:DspProcess:", "DspEstMinCount: " + this.DspEstMinCount.ToString());
                actionLog("DspProcess.cs:DspProcess:", "DispatchFromCount: " + this.DispatchFromCount.ToString());
                actionLog("DspProcess.cs:DspProcess:", "END-----------------------------------------------------------");
            }

        }
        ~DspProcess()
        {
            if (oleCnn != null)
            {
                if (oleCnn.State == System.Data.ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                oleCnn.Dispose();
            }

            if (oleSrcCnn != null)
            {
                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }

                oleSrcCnn.Dispose();
            }

            if (dbcomm != null)
            {
                dbcomm = null;
            }
        }
        public int ProcessMain()
        {
            dbcomm = new DbCommon();
            oleCnn = dbcomm.GetOleConnection("TARGET");
            oleSrcCnn = dbcomm.GetOleConnection("SOURCE");
            int li_return = 1;

            try
            {
                li_return = DspWatchDog();

                if (li_return < 0)
                {
                    actionLog("DspProcess:ProcessMain:", "DspWatchDog() return value: " + li_return.ToString());
                    return li_return;
                }

                li_return = DspMarket();

                if (li_return < 0)
                {
                    actionLog("DspProcess:ProcessMain:", "DspMarket() return value: " + li_return.ToString());
                    return li_return;
                }

                li_return = DspNewMessage();

                if (li_return < 0)
                {
                    actionLog("DspProcess:ProcessMain:", "DspNewMessage() return value: " + li_return.ToString());
                    return li_return;
                }

                li_return = DspGetSentBidNo();

                if (li_return < 0)
                {
                    actionLog("DspProcess:ProcessMain:", "DspGetSentBidNo() return value: " + li_return.ToString());
                    return li_return;
                }

                li_return = DspEmailActLog();

                if (li_return < 0)
                {
                    actionLog("DspProcess:ProcessMain:", "DspEmailActLog() return value: " + li_return.ToString());
                    return li_return;
                }

                li_return = DspIESOHeartBeat();

                if (li_return < 0)
                {
                    actionLog("DspProcess:ProcessMain:", "DspIESOHeartBeat() return value: " + li_return.ToString());
                    return li_return;
                }

                li_return = DspPLCHeartBeat();

                if (li_return < 0)
                {
                    actionLog("DspProcess:ProcessMain:", "DspPLCHeartBeat() return value: " + li_return.ToString());
                    return li_return;
                }

                li_return = DspLoadStatus();

                if (li_return < 0)
                {
                    actionLog("DspProcess:ProcessMain:", "DspLoadStatus() return value: " + li_return.ToString());
                    return li_return;
                }

                li_return = this.DspPwrOnHeartBeat();

                if (li_return < 0)
                {
                    actionLog("DspProcess:ProcessMain:", "DspPwrOnHeartBeat() return value: " + li_return.ToString());
                    return li_return;
                }
            }

            catch (Exception ex)
            {
                StackTrace st = new StackTrace(new StackFrame(true));
                actionLog("DspProcess:ProcessMain(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error in ProcessMain(): " + ex.ToString());

                li_return = -1;

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspWatchDog()
        {
            int li_return = 0;

            //1.get email address
            try
            {
                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }

                actionLog("DspProcess:DspWatchDog:", "Executing ssp_update_batch_watchdog.");
                li_return = dbcomm.updOleProData("ssp_update_batch_watchdog", oleCnn);
                actionLog("DspProcess:DspWatchDog:", "Done with ssp_update_batch_watchdog (" + li_return.ToString() + ")");

                oleCnn.Close();
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspWatchDog(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspMarket()
        {
            int li_return = 1;
            string ls_curr_or_market = "";

            ArrayList parmList = new ArrayList();
            ArrayList parmName = new ArrayList();
            ArrayList parmType = new ArrayList();
            ArrayList parmValue = new ArrayList();
            ArrayList parmSize = new ArrayList();
            DataTable dt = new DataTable();

            try
            {
                //1.get or market value from source database
                if (oleSrcCnn.State == ConnectionState.Closed)
                {
                    oleSrcCnn.Open();
                }

                actionLog("DspProcess:DspMarket:", "Executing ssp_get_or_market.");
                dt = dbcomm.getOleProData("ssp_get_or_market", oleSrcCnn);
                actionLog("DspProcess:DspMarket:", "Done with ssp_get_or_market (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        parmValue.Add(myRow.ItemArray[0]);
                        ls_curr_or_market = myRow.ItemArray[0].ToString();
                        actionLog("DspProcess:DspMarket:", "OR Market value: " + myRow.ItemArray[0].ToString());
                    }

                    if (this.IesoOrMarket != ls_curr_or_market)
                    {
                        parmName.Add("@as_or_market");
                        parmType.Add(OleDbType.Char);
                        parmSize.Add(16);

                        //2. update or market to target database
                        if (oleCnn.State == ConnectionState.Closed)
                        {
                            oleCnn.Open();
                        }

                        actionLog("DspProcess:DspMarket:", "Executing ssp_update_imo_or_market.");
                        parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                        li_return = dbcomm.updOleProData("ssp_update_imo_or_market", parmList, oleCnn);
                        actionLog("DspProcess:DspMarket:", "Done with ssp_update_imo_or_market (" + li_return.ToString() + ")");
                    }

                    this.IesoOrMarket = ls_curr_or_market;
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspMarket(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspNewMessage()
        {
            int li_return = 1;
            string ls_msg = "";

            ArrayList parmList = new ArrayList();
            ArrayList parmName = new ArrayList();
            ArrayList parmType = new ArrayList();
            ArrayList parmValue = new ArrayList();
            ArrayList parmSize = new ArrayList();
            ArrayList resultList = new ArrayList();
            DataTable dt = new DataTable();
            DataTable updDt = new DataTable();

            int li_update_stat = 0;
            string ls_response = "";

            try
            {
                //1.get New Message value from source database
                if (oleSrcCnn.State == ConnectionState.Closed)
                {
                    oleSrcCnn.Open();
                }

                actionLog("DspProcess:DspNewMessage:", "Executing ssp_poll_new_message.");
                dt = dbcomm.getOleProData("ssp_poll_new_message", oleSrcCnn);
                actionLog("DspProcess:DspNewMessage:", "Done with ssp_poll_new_message (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        foreach (DataColumn myCol in dt.Columns)
                        {
                            resultList.Add(myRow[myCol]);
                        }

                        actionLog("DspProcess:DspNewMessage:", "New Dispatch Load Message---------------------------------");
                        actionLog("DspProcess:DspNewMessage:", "Rule type: " + resultList[0].ToString());
                        actionLog("DspProcess:DspNewMessage:", "MW Amount: " + resultList[1].ToString());
                        actionLog("DspProcess:DspNewMessage:", "Message Id: " + resultList[2].ToString());
                        actionLog("DspProcess:DspNewMessage:", "Market: " + resultList[3].ToString());
                        actionLog("DspProcess:DspNewMessage:", "Disp Hour: " + resultList[4].ToString());
                        actionLog("DspProcess:DspNewMessage:", "Disp Int: " + resultList[5].ToString());
                        actionLog("DspProcess:DspNewMessage:", "Schedule Dollar: " + resultList[6].ToString());
                        actionLog("DspProcess:DspNewMessage:", "OR Market: " + resultList[7].ToString());
                        actionLog("DspProcess:DspNewMessage:", "END-------------------------------------------------------");

                        parmName.Add("@as_rule_type");
                        parmName.Add("@ai_amtmw");
                        parmName.Add("@as_market");
                        parmName.Add("@ai_disphour");
                        parmName.Add("@ai_dispint");
                        parmName.Add("@ad_sched_dollar");
                        parmName.Add("@ad_sched_mw");

                        parmType.Add(OleDbType.Char);
                        parmType.Add(OleDbType.Decimal);
                        parmType.Add(OleDbType.Char);
                        parmType.Add(OleDbType.Integer);
                        parmType.Add(OleDbType.Integer);
                        parmType.Add(OleDbType.Decimal);
                        parmType.Add(OleDbType.Decimal);

                        parmSize.Add(10);
                        parmSize.Add(5);
                        parmSize.Add(20);
                        parmSize.Add(5);
                        parmSize.Add(5);
                        parmSize.Add(8);
                        parmSize.Add(5);
                        //-----------------
                        // Gets a NumberFormatInfo associated with the en-US culture.
                        //System.Globalization.NumberFormatInfo nfi = new System.Globalization.CultureInfo( "en-US", false ).NumberFormat;

                        parmValue.Add(resultList[0]);

                        //-------------------------
                        //nfi.NumberDecimalDigits = 1;
                        parmValue.Add(System.Convert.ToDecimal(resultList[1].ToString()));
                        //-------------------------
                        parmValue.Add(resultList[3]);
                        parmValue.Add(resultList[4]);
                        parmValue.Add(resultList[5]);
                        parmValue.Add(System.Convert.ToDecimal(resultList[6].ToString()));
                        parmValue.Add(System.Convert.ToDecimal(resultList[7].ToString()));

                        //-----------------------------------------------------------------------
                        //2. update dispatch Load information into IESO instruction table in target database
                        if (updDt != null)
                        {
                            updDt.Dispose();
                        }

                        updDt = new DataTable();

                        if (oleCnn.State == ConnectionState.Closed)
                        {
                            oleCnn.Open();
                        }

                        actionLog("DspProcess:DspNewMessage:", "Executing ssp_imo_dispatch_load.");
                        parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                        updDt = dbcomm.getOleProData("ssp_imo_dispatch_load", parmList, oleCnn);
                        actionLog("DspProcess:DspNewMessage:", "Done with ssp_imo_dispatch_load (" + updDt.Rows.Count.ToString() + ")");

                        resultList = new ArrayList();

                        if (updDt.Rows.Count > 0)
                        {
                            foreach (DataRow updRow in updDt.Rows)
                            {
                                foreach (DataColumn updCol in updDt.Columns)
                                {
                                    resultList.Add(updRow[updCol]);
                                }

                                li_update_stat = Int32.Parse(resultList[0].ToString());
                                ls_response = resultList[1].ToString();

                                actionLog("DspProcess:DspNewMessage:", "Update IESO instruction----------------------------");
                                actionLog("DspProcess:DspNewMessage:", "Status: " + li_update_stat.ToString());
                                actionLog("DspProcess:DspNewMessage:", "Response: " + ls_response);
                                actionLog("DspProcess:DspNewMessage:", "MW Amt: " + parmValue[1].ToString());
                                actionLog("DspProcess:DspNewMessage:", "Disp Type: " + parmValue[3].ToString());
                                actionLog("DspProcess:DspNewMessage:", "Message Id: " + parmValue[6].ToString());
                                actionLog("DspProcess:DspNewMessage:", "END------------------------------------------------");

                                if (li_update_stat < 0)
                                {
                                    ls_msg = "Error Update IESO instruction status: " + li_update_stat.ToString() + "\r\n" +
                                                str + "Response: " + ls_response + "\r\n" +
                                                str + "MW Amt: " + parmValue[1].ToString() + "\r\n" +
                                                str + "Disp Type: " + parmValue[3].ToString() + "\r\n" +
                                                str + "Message Id: " + parmValue[6].ToString();

                                    //send email to support
                                    this.sendEmail("Error DspPatch:DspProcess:DspNewMessage:", ls_msg);
                                }
                                else
                                {
                                    // if update successfully, assign response back to IESO
                                    actionLog("DspProcess:DspNewMessage:", "Successfully Updated Instruction (" + ls_response + ")");

                                    li_return = this.DspMsgBackToIESO(ls_response);

                                    if (li_return < 0)
                                    {
                                        StackTrace st = new StackTrace(new StackFrame(true));
                                        actionLog("DspProcess:DspNewMessage(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Error calling DspMsgBackToIESO(). Return value: " + li_return.ToString());
                                        return li_return;
                                    }

                                    //2) after seedback IESO send new instruction email to user.
                                    actionLog("DspProcess:DspNewMessage:", "DspMsgBackToIESO called successfully. Sent new instruction email to user.");

                                    li_return = DspSendInstruction();

                                    if (li_return < 0)
                                    {
                                        StackTrace st = new StackTrace(new StackFrame(true));
                                        actionLog("DspProcess:DspNewMessage(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Error Calling DspSendInstruction(). Return value: " + li_return.ToString());
                                        return li_return;
                                    }

                                    actionLog("DspProcess:DspNewMessage:", "DspSendInstruction called successfully.");
                                }
                            }
                        }
                        else
                        {
                            actionLog("DspProcess:DspNewMessage:", "Executing ssp_imo_dispatch_load returned 0 rows");
                        }
                    }
                }
                else
                {
                    actionLog("DspProcess:DspNewMessage:", "Executing ssp_poll_new_message returned 0 rows");
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspNewMessage(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exeception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspSendInstruction()
        {
            int li_return = 1;
            int li_emailsent = 0;
            string ls_subject = "";
            string ls_emailbody = "";
            string ls_msg = "";
            string ls_instype = "";
            string ls_imo_instruction = "";
            string ls_market = "";
            string ls_eng_mw = "";

            DataTable dt = new DataTable();
            ArrayList parmList = new ArrayList();
            ArrayList parmTypeList = new ArrayList();
            ArrayList parmValueList = new ArrayList();
            ArrayList parmSizeList = new ArrayList();
            ArrayList resultValue = new ArrayList();

            //1.get email address
            try
            {
                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }

                actionLog("DspProcess:DspSendInstruction:", "Executing ssp_batch_get_imo_instruction.");
                dt = dbcomm.getOleProData("ssp_batch_get_imo_instruction", oleCnn);
                actionLog("DspProcess:DspSendInstruction:", "Done with ssp_batch_get_imo_instruction (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        resultValue = null;
                        parmList = null;
                        parmTypeList = null;
                        parmValueList = null;
                        parmList = new ArrayList();
                        parmTypeList = new ArrayList();
                        parmValueList = new ArrayList();
                        parmSizeList = new ArrayList();
                        resultValue = new ArrayList();

                        foreach (DataColumn myCol in dt.Columns)
                        {
                            resultValue.Add(myRow[myCol]);
                        }

                        ls_instype = resultValue[1].ToString().ToUpper();
                        ls_imo_instruction = resultValue[5].ToString().ToUpper();
                        ls_market = resultValue[3].ToString().ToUpper();
                        ls_eng_mw = resultValue[2].ToString().ToUpper();

                        actionLog("DspProcess:DspSendInstruction:", "Get instruction information--------------------------------------");
                        actionLog("DspProcess:DspSendInstruction:", "seq_no: " + resultValue[0].ToString());
                        actionLog("DspProcess:DspSendInstruction:", "instype: " + ls_instype);
                        actionLog("DspProcess:DspSendInstruction:", "megwat: " + resultValue[2].ToString());
                        actionLog("DspProcess:DspSendInstruction:", "market: " + ls_market);
                        actionLog("DspProcess:DspSendInstruction:", "approval Ind: " + resultValue[4].ToString());
                        actionLog("DspProcess:DspSendInstruction:", "instruction:" + ls_imo_instruction);
                        actionLog("DspProcess:DspSendInstruction:", "trafficlight: " + resultValue[6].ToString());
                        actionLog("DspProcess:DspSendInstruction:", "updatestatus: " + resultValue[7].ToString());
                        actionLog("DspProcess:DspSendInstruction:", "rejectaccept: " + resultValue[8].ToString());
                        actionLog("DspProcess:DspSendInstruction:", "END--------------------------------------------------------------");

                        ls_imo_instruction = ls_imo_instruction.Trim();
                        ls_instype = ls_instype.Trim();
                        ls_market = ls_market.Trim();

                        actionLog("DspProcess:DspSendInstruction:", "Start call DSpDBActionLog(" + ls_instype + ", " + ls_imo_instruction + ", " + ls_market + ")");

                        this.DspDbActionLog(ls_instype, ls_imo_instruction, ls_market);

                        parmList.Add("@dsp_type");
                        parmList.Add("@scenario");

                        parmTypeList.Add(OleDbType.Char);
                        parmTypeList.Add(OleDbType.Char);

                        parmSizeList.Add(16);
                        parmSizeList.Add(16);

                        //1. ORActive
                        if ((ls_instype == "ORA") & (ls_imo_instruction == "ORACTIVATION"))
                        {
                            parmValueList.Add("INSTRUCTION");
                            parmValueList.Add("INSTRUCTION");

                            ls_subject = "New ORActivation instruction";
                            ls_emailbody = ls_oractive + " " + ls_eng_mw + "MW";

                            actionLog("DspProcess:DspSendInstruction:", "Calling sendEamil method for new instruction ORACTIVATION.");
                            li_emailsent = sendEmail("ssp_get_dsp_email_addr", parmList, parmTypeList, parmSizeList, parmValueList, ls_subject, ls_emailbody);

                            if (li_emailsent >= 0)
                            {
                                StackTrace st = new StackTrace(new StackFrame(true));
                                actionLog("DspProcess:DspSendInstruction(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Email has been sent successfully");
                            }
                            else
                            {
                                li_emailsent = sendEmail(ls_subject, ls_emailbody);
                                StackTrace st = new StackTrace(new StackFrame(true));
                                actionLog("DspProcess:DspSendInstruction(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Email has been sent to the default address");
                            }

                            ls_msg = "Dispatch Instruction: " + ls_instype + ", " + ls_imo_instruction + " (" + ls_oractive + ")";
                            actionLog("DspProcess:DspSendInstruction:", ls_msg);
                        }

                        //2. Dispatch Off
                        if ((ls_instype == "ENG") & (ls_imo_instruction == "DISPATCHOFF"))
                        {
                            parmValueList.Add("INSTRUCTION");
                            parmValueList.Add("INSTRUCTION");

                            ls_subject = "New Dispatch Off instruction";
                            ls_emailbody = ls_engoff + " " + ls_eng_mw + "MW";

                            actionLog("DspProcess:DspSendInstruction:", "Calling sendEamil method for new instruction DISPATCHOFF.");
                            li_emailsent = sendEmail("ssp_get_dsp_email_addr", parmList, parmTypeList, parmSizeList, parmValueList, ls_subject, ls_emailbody);

                            if (li_emailsent >= 0)
                            {
                                StackTrace st = new StackTrace(new StackFrame(true));
                                actionLog("DspProcess:DspSendInstruction(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Email has been sent successfully");
                            }
                            else
                            {
                                li_emailsent = sendEmail(ls_subject, ls_emailbody);
                                StackTrace st = new StackTrace(new StackFrame(true));
                                actionLog("DspProcess:DspSendInstruction(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Email has been sent to the default address");
                            }

                            ls_msg = "Dispatch Instruction: " + ls_instype + ", " + ls_imo_instruction + " (" + ls_engoff + ")";
                            actionLog("DspProcess:DspSendInstruction:", ls_msg);
                        }

                        //3. Dispatch On
                        if ((ls_instype == "ENG") & (ls_imo_instruction == "DISPATCHON"))
                        {
                            parmValueList.Add("INSTRUCTION");
                            parmValueList.Add("INSTRUCTION");

                            ls_subject = "New Dispatch On instruction";
                            ls_emailbody = ls_engon + " " + ls_eng_mw + "MW";

                            actionLog("DspProcess:DspSendInstruction:", "Calling sendEamil method for new instruction DISPATCHON.");
                            li_emailsent = sendEmail("ssp_get_dsp_email_addr", parmList, parmTypeList, parmSizeList, parmValueList, ls_subject, ls_emailbody);

                            if (li_emailsent >= 0)
                            {
                                StackTrace st = new StackTrace(new StackFrame(true));
                                actionLog("DspProcess:DspSendInstruction(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Email has been sent successfully");
                            }
                            else
                            {
                                li_emailsent = sendEmail(ls_subject, ls_emailbody);

                                StackTrace st = new StackTrace(new StackFrame(true));
                                actionLog("DspProcess:DspSendInstruction(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Email has been sent to the default address");
                            }

                            ls_msg = "Dispatch Instruction: " + ls_instype + ", " + ls_imo_instruction + " (" + ls_engon + ")";
                            actionLog("DspProcess:DspSendInstruction:", ls_msg);
                        }
                    }
                }
                else
                {
                    actionLog("DspProcess:DspSendInstruction:", "Executing ssp_batch_get_imo_instruction returned 0 rows");
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspSendInstruction(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspDbActionLog(string ps_instype, string ps_instruction, string ps_market)
        {
            int li_return = 1;
            string ls_act_type = "INSTRUCTION";
            string ls_act_desc = "";
            string ls_comp_name = "";
            string ls_login = "";

            ArrayList parmName = new ArrayList();
            ArrayList parmType = new ArrayList();
            ArrayList parmValue = new ArrayList();
            ArrayList parmSize = new ArrayList();
            ArrayList parmList = new ArrayList();

            try
            {
                switch (ps_instype.ToUpper())
                {
                    case "ENG":
                        if (ps_instruction == "DISPATCHOFF")
                        {
                            ls_act_desc = "NEW DISPATCH OFF INSTRUCTION";
                        }
                        else if (ps_instruction == "DISPATCHON")
                        {
                            ls_act_desc = "NEW DISPATCH ON INSTRUCTION";
                        }
                        else
                        {
                            ls_act_desc = "NEW UNKNOWN ENERGY INSTRUCTION";
                        }

                        break;

                    case "ORA":
                        ls_act_desc = "NEW ORACTIVATION INSTRUCTION";
                        break;

                    case "RESV":
                        if (ps_market == "30R")
                        {
                            ls_act_desc = "NEW 30R RESERVE INSTRUCTION";
                        }
                        else if (ps_market == "10N")
                        {
                            ls_act_desc = "NEW 10N RESERVE INSTRUCTION";
                        }
                        else
                        {
                            ls_act_desc = "NEW UNKNOWN RESERVE INSTRUCTION";
                        }
                        break;

                    default:
                        ls_act_desc = "NEW UNKNOWN INSTRUCTION";
                        break;
                }

                li_return = getComputerName(ref ls_comp_name, ref ls_login);

                if (li_return == -1)
                {
                    ls_comp_name = "none";
                    ls_login = "none";

                    this.actionLog("DspProcess:DspDbActionLog:", "Could Not Establish Computer Name");
                }

                parmName.Add("@act_type");
                parmName.Add("@act_desc");
                parmName.Add("@act_computer");
                parmName.Add("@act_comp_login");
                parmName.Add("@act_status");

                parmValue.Add(ls_act_type);
                parmValue.Add(ls_act_desc);
                parmValue.Add(ls_comp_name);
                parmValue.Add(ls_login);
                parmValue.Add("S");

                parmType.Add(OleDbType.Char);
                parmType.Add(OleDbType.Char);
                parmType.Add(OleDbType.Char);
                parmType.Add(OleDbType.Char);
                parmType.Add(OleDbType.Char);

                parmSize.Add(16);
                parmSize.Add(30);
                parmSize.Add(30);
                parmSize.Add(16);
                parmSize.Add(16);

                //2. update or market to target database
                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }
                actionLog("DspProcess:DspDbActionLog:", "Executing exec ssp_insert_dsp_act_log.");
                parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                li_return = dbcomm.updOleProData("ssp_insert_dsp_act_log", parmList, oleCnn);
                actionLog("DspProcess:DspDbActionLog:", "Done with ssp_insert_dsp_act_log (" + li_return.ToString() + ")");
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspDbActionLog(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspDbActionLog(string ps_actType, string ps_actDesc)
        {
            int li_return = 1;
            string ls_comp_name = "";
            string ls_login = "";

            ArrayList parmName = new ArrayList();
            ArrayList parmType = new ArrayList();
            ArrayList parmValue = new ArrayList();
            ArrayList parmSize = new ArrayList();
            ArrayList parmList = new ArrayList();

            try
            {
                li_return = getComputerName(ref ls_comp_name, ref ls_login); if (li_return == -1)
                {
                    ls_comp_name = "none";
                    ls_login = "none";

                    this.actionLog("DspProcess:DspDbActionLog:", "Could Not Establish Computer Name");
                }

                parmName.Add("@act_type");
                parmName.Add("@act_desc");
                parmName.Add("@act_computer");
                parmName.Add("@act_comp_login");
                parmName.Add("@act_status");

                parmValue.Add(ps_actType);
                parmValue.Add(ps_actDesc);
                parmValue.Add(ls_comp_name);
                parmValue.Add(ls_login);
                parmValue.Add("S");

                parmType.Add(OleDbType.Char);
                parmType.Add(OleDbType.Char);
                parmType.Add(OleDbType.Char);
                parmType.Add(OleDbType.Char);
                parmType.Add(OleDbType.Char);

                parmSize.Add(16);
                parmSize.Add(30);
                parmSize.Add(30);
                parmSize.Add(16);
                parmSize.Add(16);

                //2. update or market to target database
                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }
                actionLog("DspProcess:DspDbActionLog:", "Executing ssp_insert_dsp_act_log.");
                parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                li_return = dbcomm.updOleProData("ssp_insert_dsp_act_log", parmList, oleCnn);
                actionLog("DspProcess:DspDbActionLog:", "Done with ssp_insert_dsp_act_log (" + li_return.ToString() + ")");
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspDbActionLog(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspMsgBackToIESO(string ps_response)
        {
            int li_return = 1;

            ArrayList parmList = new ArrayList();
            ArrayList parmName = new ArrayList();
            ArrayList parmType = new ArrayList();
            ArrayList parmValue = new ArrayList();
            ArrayList parmSize = new ArrayList();

            try
            {
                //1.feed back msg status to IESO source database
                if (oleSrcCnn.State == ConnectionState.Closed)
                {
                    oleSrcCnn.Open();
                }

                parmName.Add("@as_response");
                parmType.Add(OleDbType.Char);

                parmValue.Add(ps_response);
                parmSize.Add(1);
                parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);

                actionLog("DspProcess:DspDbActionLog:", "Executing ssp_update_status_flags");
                li_return = dbcomm.updOleProData("ssp_update_status_flags", parmList, oleSrcCnn);
                actionLog("DspProcess:DspDbActionLog:", "Done with ssp_update_status_flags (" + li_return.ToString(), ")");
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspDbActionLog(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspGetBidRejectDetail(ref string rjt_reason, string bid_no)
        {
            string ls_rjt_reason1 = "";
            string ls_rjt_reason2 = "";
            string ls_rjt_reason3 = "";
            string ls_rjt_reason4 = "";
            string ls_rjt_reason5 = "";
            string ls_rjt_reason6 = "";
            string ls_rjt_reason7 = "";
            string ls_rjt_reason8 = "";
            string ls_rjt_reason9 = "";
            string ls_rjt_reason10 = "";
            int li_return = 1;

            DataTable dt;
            dt = new DataTable();

            ArrayList parmList = new ArrayList();
            ArrayList parmNameList = new ArrayList();
            ArrayList parmTypeList = new ArrayList();
            ArrayList parmValueList = new ArrayList();
            ArrayList parmSizeList = new ArrayList();
            ArrayList resultList = new ArrayList();

            parmNameList.Add("@bid_no");
            parmValueList.Add(bid_no);
            parmTypeList.Add(OleDbType.VarChar);
            parmSizeList.Add(30);

            try
            {
                if (oleSrcCnn.State == ConnectionState.Closed)
                {
                    oleSrcCnn.Open();
                }
                parmList = dbcomm.MakeInParm(parmNameList, parmTypeList, parmSizeList, parmValueList);

                actionLog("DspProcess:DspGetBidRejectDetail:", "Executing ssp_get_bid_reject_detail with bid no: " + bid_no);
                dt = dbcomm.getOleProData("ssp_get_bid_reject_detail", parmList, oleSrcCnn);
                actionLog("DspProcess:DspGetBidRejectDetail:", "Done with ssp_get_bid_reject_detail (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        foreach (DataColumn myCol in dt.Columns)
                        {
                            resultList.Add(myRow[myCol]);
                        }

                        ls_rjt_reason1 = (string)resultList[0];
                        ls_rjt_reason2 = (string)resultList[1];
                        ls_rjt_reason3 = (string)resultList[2];
                        ls_rjt_reason4 = (string)resultList[3];
                        ls_rjt_reason5 = (string)resultList[4];
                        ls_rjt_reason6 = (string)resultList[5];
                        ls_rjt_reason7 = (string)resultList[6];
                        ls_rjt_reason8 = (string)resultList[7];
                        ls_rjt_reason9 = (string)resultList[8];
                        ls_rjt_reason10 = (string)resultList[9];

                        if (ls_rjt_reason1 == null) { ls_rjt_reason1 = ""; }
                        if (ls_rjt_reason2 == null) { ls_rjt_reason2 = ""; }
                        if (ls_rjt_reason3 == null) { ls_rjt_reason3 = ""; }
                        if (ls_rjt_reason4 == null) { ls_rjt_reason4 = ""; }
                        if (ls_rjt_reason5 == null) { ls_rjt_reason5 = ""; }
                        if (ls_rjt_reason6 == null) { ls_rjt_reason6 = ""; }
                        if (ls_rjt_reason7 == null) { ls_rjt_reason7 = ""; }
                        if (ls_rjt_reason8 == null) { ls_rjt_reason8 = ""; }
                        if (ls_rjt_reason9 == null) { ls_rjt_reason9 = ""; }
                        if (ls_rjt_reason10 == null) { ls_rjt_reason10 = ""; }
                    }

                    rjt_reason = ls_rjt_reason1 + "\r\n" +
                        ls_rjt_reason2 + "\r\n" +
                        ls_rjt_reason3 + "\r\n" +
                        ls_rjt_reason4 + "\r\n" +
                        ls_rjt_reason5 + "\r\n" +
                        ls_rjt_reason6 + "\r\n" +
                        ls_rjt_reason7 + "\r\n" +
                        ls_rjt_reason8 + "\r\n" +
                        ls_rjt_reason9 + "\r\n" +
                        ls_rjt_reason10;
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspGetBidRejectDetail(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error (" + bid_no + "): " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspGetSentBidNo()
        {
            string ls_bidno = "";
            string ls_msg = "";
            string ls_emailSubject = "";
            string ls_emailBody = "";
            string ls_emailBodyTemp = "";
            int li_emailSent = 0;
            int li_tempBody = 0;
            int li_return = 1;

            ArrayList parmList = new ArrayList();
            ArrayList parmName = new ArrayList();
            ArrayList parmType = new ArrayList();
            ArrayList parmValue = new ArrayList();
            ArrayList parmSize = new ArrayList();
            ArrayList resultList = new ArrayList();
            DataTable dt = new DataTable();
            DataTable bidDt = new DataTable();

            try
            {
                //1.get bid no from target database
                parmName.Add("@accepted_ind");
                parmType.Add(OleDbType.Char);

                parmValue.Add("'0'");
                parmSize.Add(16);

                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }

                actionLog("DspProcess:DspGetSentBidNo:", "Executing ssp_get_sent_imo_bidno.");
                parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                dt = dbcomm.getOleProData("ssp_get_sent_imo_bidno", parmList, oleCnn);
                actionLog("DspProcess:DspGetSentBidNo:", "Done with ssp_get_sent_imo_bidno (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ls_msg = "Retrieved sent_bidno:" + myRow.ItemArray[0].ToString();
                        actionLog("DspProcess:DspGetSentBidNo:", ls_msg);

                        parmName = null;
                        parmType = null;
                        parmSize = null;
                        parmValue = null;

                        parmName = new ArrayList();
                        parmType = new ArrayList();
                        parmSize = new ArrayList();
                        parmValue = new ArrayList();

                        parmName.Add("@bidno");
                        parmType.Add(OleDbType.Char);
                        parmValue.Add(myRow.ItemArray[0]);
                        parmSize.Add(16);

                        if (bidDt != null)
                        {
                            bidDt = null;
                        }

                        bidDt = new DataTable();

                        //2.get bid status from source database
                        actionLog("DspProcess:DspGetSentBidNo:", "Executing ssp_get_bid_status.");
                        parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                        bidDt = dbcomm.getOleProData("ssp_get_bid_status", parmList, oleSrcCnn);
                        actionLog("DspProcess:DspGetSentBidNo:", "Done With ssp_get_bid_status (" + bidDt.Rows.Count.ToString() + ")");

                        if (bidDt.Rows.Count > 0)
                        {
                            foreach (DataRow bidRow in bidDt.Rows)
                            {
                                if (bidRow.ItemArray[0].ToString() != null & bidRow.ItemArray[0].ToString() != "")
                                {
                                    int li_bid_accept_ind = Int32.Parse(bidRow.ItemArray[0].ToString());

                                    if (li_bid_accept_ind.ToString() == null)
                                    {
                                        li_bid_accept_ind = 0;
                                    }

                                    actionLog("DspProcess:DspGetSentBidNo:", "Bidno: " + myRow.ItemArray[0].ToString() + ", Accepted Ind: " + li_bid_accept_ind.ToString());

                                    parmName = null;
                                    parmType = null;
                                    parmSize = null;
                                    parmValue = null;
                                    parmName = new ArrayList();
                                    parmType = new ArrayList();
                                    parmSize = new ArrayList();
                                    parmValue = new ArrayList();

                                    parmName.Add("@bidno");
                                    parmType.Add(OleDbType.Char);
                                    parmValue.Add(myRow.ItemArray[0]);
                                    parmSize.Add(16);

                                    parmName.Add("@sent_ind");
                                    parmType.Add(OleDbType.Char);
                                    parmValue.Add("Y");
                                    parmSize.Add(16);

                                    parmName.Add("@accepted_ind");
                                    parmType.Add(OleDbType.Integer);
                                    parmValue.Add(bidRow.ItemArray[0]);
                                    parmSize.Add(10);

                                    //3. update sent bid status to target database
                                    if (oleCnn.State == ConnectionState.Closed)
                                    {
                                        oleCnn.Open();
                                    }
                                    actionLog("DspProcess:DspGetSentBidNo:", "Executing ssp_update_dsp_bid_status.");
                                    parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                                    li_return = dbcomm.updOleProData("ssp_update_dsp_bid_status", parmList, oleCnn);
                                    actionLog("DspProcess:DspGetSentBidNo:", "Done with ssp_update_dsp_bid_status (" + li_return.ToString() + ")");

                                    //4.send email if the accept indicator show minus value.
                                    ls_emailSubject = "Bid Rejected";

                                    if (li_bid_accept_ind == -2)
                                    {
                                        ls_emailSubject = "Bid Partially Rejected";
                                    }

                                    ls_emailBody = "IESO has rejected the bid we sent (" + myRow.ItemArray[0].ToString() + "). Please check bid history report for details" + "\r\n";
                                    ls_emailBodyTemp = "";
                                    ls_bidno = myRow.ItemArray[0].ToString();

                                    if (li_bid_accept_ind < 0)
                                    {
                                        li_tempBody = DspGetBidRejectDetail(ref ls_emailBodyTemp, ls_bidno);

                                        if (li_tempBody < 0)
                                        {
                                            li_return = -1;
                                        }

                                        if (ls_emailBodyTemp == null)
                                        {
                                            ls_emailBodyTemp = "";
                                        }

                                        ls_emailBody = ls_emailBody + "\r\n" + ls_emailBodyTemp;
                                        li_emailSent = sendDbEmail("BID", "REJECT", ls_emailSubject, ls_emailBody);

                                        if (li_emailSent < 0)
                                        {
                                            li_return = -1;
                                        }
                                    }
                                }
                                else
                                {
                                    StackTrace st = new StackTrace(new StackFrame(true));
                                    actionLog("DspProcess:DspGetSentBidNo(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Could not find bid no. " + myRow.ItemArray[0].ToString());
                                }
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspGetSentBidNo(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (bidDt != null)
                {
                    bidDt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspEmailActLog()
        {
            int li_return = 1;
            int li_emailSent = 0;
            string ls_emailSubject = "";
            string ls_emailBody = "";
            string ls_emailType = "";
            string ls_emailSubType = "";

            ArrayList parmList = new ArrayList();
            ArrayList parmName = new ArrayList();
            ArrayList parmType = new ArrayList();
            ArrayList parmValue = new ArrayList();
            ArrayList parmSize = new ArrayList();
            ArrayList resultList = new ArrayList();
            DataTable dt = new DataTable();

            try
            {
                //1.get bid no from target database
                parmName.Add("@email_status");
                parmType.Add(OleDbType.Char);

                parmValue.Add("N");
                parmSize.Add(16);

                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }

                actionLog("DspProcess:DspEmailActLog:", "Executing ssp_get_dsp_email_act_status.");
                parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                dt = dbcomm.getOleProData("ssp_get_dsp_email_act_status", parmList, oleCnn);
                actionLog("DspProcess:DspEmailActLog:", "Done with ssp_get_dsp_email_act_status (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        resultList = null;
                        resultList = new ArrayList();

                        foreach (DataColumn myCol in dt.Columns)
                        {
                            resultList.Add(myRow[myCol]);
                        }

                        actionLog("DspProcess:DspEmailActLog:", "Email Action Log Information----------------------------------------------");
                        actionLog("DspProcess:DspEmailActLog:", "Email Act ID:" + resultList[0].ToString());
                        actionLog("DspProcess:DspEmailActLog:", "Email Act Type:" + resultList[1].ToString());
                        actionLog("DspProcess:DspEmailActLog:", "Email Act Sub Type:" + resultList[2].ToString());
                        actionLog("DspProcess:DspEmailActLog:", "Email Act Subject:" + resultList[3].ToString());
                        actionLog("DspProcess:DspEmailActLog:", "Email Act Content:" + resultList[4].ToString());
                        actionLog("DspProcess:DspEmailActLog:", "Email Act Name:" + resultList[5].ToString());
                        actionLog("DspProcess:DspEmailActLog:", "Email Act Login:" + resultList[6].ToString());
                        actionLog("DspProcess:DspEmailActLog:", "Email Act Status:" + resultList[7].ToString());
                        actionLog("DspProcess:DspEmailActLog:", "END-----------------------------------------------------------------------");

                        //2. send email and update email act status
                        ls_emailType = resultList[1].ToString();

                        if (ls_emailType == "CONTROLOFF")
                        {
                            ls_emailType = "DSPCNTLOFF";
                        }

                        ls_emailSubType = resultList[2].ToString();

                        if (ls_emailSubType == "CONTROLOFF")
                        {
                            ls_emailSubType = "DSPCNTLOFF";
                        }

                        ls_emailSubject = resultList[3].ToString();
                        ls_emailBody = resultList[4].ToString();

                        li_emailSent = sendDbEmail(ls_emailType, ls_emailSubType, ls_emailSubject, ls_emailBody);

                        if (li_emailSent < 0)
                        {
                            li_return = -1;
                            break;
                        }

                        //3. update email act log status.
                        parmName = null;
                        parmType = null;
                        parmSize = null;
                        parmValue = null;

                        parmName = new ArrayList();
                        parmType = new ArrayList();
                        parmSize = new ArrayList();
                        parmValue = new ArrayList();

                        parmName.Add("@email_id");
                        parmType.Add(OleDbType.Integer);
                        parmValue.Add(resultList[0]);
                        parmSize.Add(32);

                        parmName.Add("@email_sent_ind");
                        parmType.Add(OleDbType.Char);
                        parmValue.Add("Y");
                        parmSize.Add(16);

                        if (oleCnn.State == ConnectionState.Closed)
                        {
                            oleCnn.Open();
                        }

                        actionLog("DspProcess:DspEmailActLog", "Executing ssp_upd_dsp_email_act_status.");
                        parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                        li_return = dbcomm.updOleProData("ssp_upd_dsp_email_act_status", parmList, oleCnn);
                        actionLog("DspProcess:DspEmailActLog", "Done with ssp_upd_dsp_email_act_status (" + li_return.ToString() + ")");
                    }
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspEmailActLog(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspPwrOnHeartBeat()
        {
            int li_return = 1;
            int li_emailSent = 0;
            string ls_emailSubject = "";
            string ls_emailBody = "";
            string ls_msg = "";
            string ls_noaction_ind = "";

            ArrayList resultList = new ArrayList();
            DataTable dt = new DataTable();

            try
            {
                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }

                actionLog("DspProcess:DspPwrOnHeartBeat:", "Executing ssp_dsp_poweron_monitor.");
                dt = dbcomm.getOleProData("ssp_dsp_poweron_monitor", oleCnn);
                actionLog("DspProcess:DspPwrOnHeartBeat:", "Done With ssp_dsp_poweron_monitor (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        foreach (DataColumn myCol in dt.Columns)
                        {
                            resultList.Add(myRow[myCol]);
                        }

                        ls_noaction_ind = resultList[0].ToString();

                        if (ls_noaction_ind.ToUpper() == "Y")
                        {
                            ls_msg = "Power On Update Not Active for over 2 minutes. System unixtime: " + resultList[1].ToString() + "; Upd db unixtime: " + resultList[2].ToString();
                            actionLog("DspProcess:DspPwrOnHeartBeat:", ls_msg);

                            ls_emailSubject = "Power On Update Not Active";
                            ls_emailBody = ls_msg;

                            if (this.PwrOnBeatMailSent == "N")
                            {
                                li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);
                                actionLog("DspProcess:DspPwrOnHeartBeat:", "Email send stauts: " + li_emailSent.ToString());

                                if (li_emailSent < 0)
                                {
                                    li_return = -1;
                                    break;
                                }
                                else
                                {
                                    //assign email sent indicator
                                    this.PwrOnBeatMailSent = "Y";
                                    actionLog("DspProcess:DspPwrOnHeartBeat:", "PwrOnBeatMailSent Assigned To \"Y\"");
                                }

                                //3. update target database action log table
                                li_return = DspDbActionLog("PLCHEARTBEAT", ls_emailBody);

                                if (li_return < 0)
                                {
                                    break;
                                }
                            }
                            else
                            {
                                //3. update target database action log table
                                ls_emailBody = "Power On update heartbeat has been sent out to user to wait to fix FLINK or DB problem.";

                                if (li_return < 0)
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            //check whether the IESO heart beat release or not.
                            if (this.PwrOnBeatMailSent == "Y")
                            {
                                ls_msg = "Power On Update Heartbeat restored at: " + getStndSummerDateTime().ToString() +
                                            ". System unixtime: " + resultList[1].ToString() +
                                            "; Upd db unixtime: " + resultList[2].ToString();

                                actionLog("DspProcess:DspPwrOnHeartBeat:", ls_msg);

                                ls_emailSubject = "Power On Update Active has been restored";
                                ls_emailBody = ls_msg;

                                li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                                if (li_emailSent < 0)
                                {
                                    li_return = -1;
                                    break;
                                }
                                else
                                {
                                    this.PwrOnBeatMailSent = "N";
                                }
                            }
                            else
                            {
                                ls_msg = "Power On Update Active at: " + getStndSummerDateTime().ToString() +
                                            ". System unixtime: " + resultList[1].ToString() +
                                            "; Upd db unixtime: " + resultList[2].ToString();

                                actionLog("DspProcess:DspPwrOnHeartBeat:", ls_msg);
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspPwrOnHeartBeat(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspIESOHeartBeat()
        {
            int li_return = 1;
            int li_emailSent = 0;
            string ls_emailSubject = "";
            string ls_emailBody = "";
            string ls_msg = "";
            string ls_noaction_ind = "";

            Int32 li_iesounixtime = 0;
            ArrayList resultList = new ArrayList();
            DataTable dt = new DataTable();

            try
            {
                if (oleSrcCnn.State == ConnectionState.Closed)
                {
                    oleSrcCnn.Open();
                }

                actionLog("DspProcess:DspIESOHeartBeat:", "Executing ssp_check_sys_no_action.");
                dt = dbcomm.getOleProData("ssp_check_sys_no_action", oleSrcCnn);
                actionLog("DspProcess:DspIESOHeartBeat:", "Done with ssp_check_sys_no_action (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        foreach (DataColumn myCol in dt.Columns)
                        {
                            resultList.Add(myRow[myCol]);
                        }

                        try
                        {
                            li_iesounixtime = Int32.Parse(resultList[2].ToString());

                            if (li_iesounixtime != this.IESOUnixTime)
                            {
                                this.IESOUnixDateTime = getStndSummerDateTime();
                            }
                        }

                        catch (Exception ex)
                        {
                            StackTrace st = new StackTrace(new StackFrame(true));
                            actionLog("DspProcess:DspIESOHeartBeat(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Unix Time Discrepency: " + ex.ToString());
                        }

                        ls_noaction_ind = resultList[0].ToString();

                        if (ls_noaction_ind.ToUpper() == "Y")
                        {
                            ls_msg = "IESO HeartBeat Not Active at: " + this.IESOUnixDateTime.ToString() +
                                        ". System unixtime: " + resultList[1].ToString() +
                                        "; IESO db unixtime: " + resultList[2].ToString();

                            actionLog("DspProcess:DspIESOHeartBeat:", ls_msg);

                            ls_emailSubject = "IESO HeartBeat Not Active";
                            ls_emailBody = ls_msg;

                            if (this.IesoBeatMailSent == "N")
                            {
                                li_emailSent = sendDbEmail("IMOHEARTBEAT", "IMOHEARTBEAT", ls_emailSubject, ls_emailBody);
                                actionLog("DspProcess:DspIESOHeartBeat:", "Email send stauts value: " + li_emailSent.ToString());

                                if (li_emailSent < 0)
                                {
                                    li_return = -1;
                                    break;
                                }
                                else
                                {
                                    this.IesoBeatMailSent = "Y";
                                    actionLog("DspProcess:DspIESOHeartBeat:", "IESOBeatMailSent assigned to \"Y\".");
                                }

                                li_return = DspDbActionLog("IMOHEARTBEAT", ls_emailBody);

                                if (li_return < 0)
                                {
                                    break;
                                }
                            }
                            else
                            {
                                ls_emailBody = "IESO heart beat has been sent out to user to wait to fix IESO heartbeat problem.";

                                if (li_return < 0)
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            //check whether the IESO heart beat release or not.
                            if (this.IesoBeatMailSent == "Y")
                            {
                                ls_msg = "IESO Heartbeat restored at: " + getStndSummerDateTime().ToString() +
                                            ". System unixtime: " + resultList[1].ToString() +
                                            "; IESO db unixtime: " + resultList[2].ToString();

                                actionLog("DspProcess:DspIESOHeartBeat:", ls_msg);

                                ls_emailSubject = "IESO HeartBeat Active has been restored";
                                ls_emailBody = ls_msg;

                                li_emailSent = sendDbEmail("IMOHEARTBEAT", "IMOHEARTBEAT", ls_emailSubject, ls_emailBody);

                                if (li_emailSent < 0)
                                {
                                    li_return = -1;
                                    break;
                                }
                                else
                                {
                                    this.IesoBeatMailSent = "N";
                                }
                            }
                            else
                            {
                                ls_msg = "IESO Active at: " + getStndSummerDateTime().ToString() +
                                            ". System unixtime: " + resultList[1].ToString() +
                                            "; IESO db unixtime: " + resultList[2].ToString();
                                actionLog("DspProcess:DspIESOHeartBeat:", ls_msg);
                            }
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspIESOHeartBeat(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspPLCHeartBeat()
        {
            int li_return = 1;
            int li_emailSent = 0;
            int li_plccounterwr = 0;
            int li_plccounterrd = 0;
            int li_timespan_curr = 0;
            string ls_emailSubject = "";
            string ls_emailBody = "";
            string ls_msg = "";

            ArrayList parmList = new ArrayList();
            ArrayList parmName = new ArrayList();
            ArrayList parmType = new ArrayList();
            ArrayList parmValue = new ArrayList();
            ArrayList parmSize = new ArrayList();
            ArrayList resultList = new ArrayList();
            DataTable dt = new DataTable();

            //assing istance value is_plcbeat_emailsent ='Y'
            //check the time span>2 sent email plc heart beat
            //check time span <2 and check istance plc heat beat email sent indicator,
            //if "Y" then send email say PLC heat has been restored. else log to file.
            try
            {
                //1.get PLC counter value from target database
                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }

                actionLog("DspProcess:DspPLCHeartBeat:", "Executing ssp_get_plc_heartbeat_count.");
                dt = dbcomm.getOleProData("ssp_get_plc_heartbeat_count", oleCnn);
                actionLog("DspProcess:DspPLCHeartBeat:", "Done with ssp_get_plc_heartbeat_count (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        foreach (DataColumn myCol in dt.Columns)
                        {
                            resultList.Add(myRow[myCol]);
                        }

                        li_plccounterrd = Int32.Parse(resultList[0].ToString());
                        li_plccounterwr = Int32.Parse(resultList[1].ToString());


                        ls_msg = "PlcCounter Read: " + li_plccounterrd.ToString() +
                                        " Write: " + li_plccounterwr.ToString() +
                                        " Write Previous: " + this.PlcCntWrPrev.ToString() +
                                        " Current Date: " + getStndSummerDateTime().ToString() +
                                        " Previous Write Date: " + this.PlcCntWrDatePrev.ToString();

                        actionLog("DspProcess:DspPLCHeartBeat:", ls_msg);

                        ls_emailSubject = "PLC HeartBeat Not Active";
                        ls_emailBody = ls_msg;

                        if (this.PlcCntWrPrev != li_plccounterwr)
                        {
                            actionLog("DspProcess:DspPLCHeartBeat:", "PlcCounter values; write: " + li_plccounterwr.ToString() +
                                            " Previous Write: " + this.PlcCntWrPrev.ToString() +
                                            " Current Date: " + getStndSummerDateTime().ToString() +
                                            " Previous Write Date: " + this.PlcCntWrDatePrev.ToString());
                            this.PlcCntWrDatePrev = getStndSummerDateTime();
                        }

                        if (this.PlcCntWrPrev == li_plccounterwr)
                        {
                            TimeSpan span = (getStndSummerDateTime()).Subtract(this.PlcCntWrDatePrev);
                            li_timespan_curr = (int)span.TotalMinutes;

                            actionLog("DspProcess:DspPLCHeartBeat:", "PLC heartbeat time span in minutes: " + li_timespan_curr.ToString());

                            int li_beat_timeout = Int32.Parse(Panel.strPlcHeartBeatMinute);

                            if (li_timespan_curr >= li_beat_timeout)
                            {
                                if (this.PlcBeatMailSent == "N")
                                {
                                    ls_msg = "IESO PLC HeartBeat Not Active at: " + this.PlcCntWrDatePrev.ToString() +
                                                ". PlcCounterRd : " + resultList[0].ToString() +
                                                "; PlcCounterWr : " + resultList[1].ToString();

                                    actionLog("DspProcess:DspPLCHeartBeat:", ls_msg);

                                    ls_emailSubject = "PLC HeartBeat Not Active";
                                    ls_emailBody = ls_msg;

                                    li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                                    if (li_emailSent < 0)
                                    {
                                        li_return = -1;
                                        break;
                                    }
                                    else
                                    {
                                        this.PlcBeatMailSent = "Y";
                                    }

                                    li_return = DspDbActionLog("PLCHEARTBEAT", ls_emailBody);

                                    if (li_return < 0)
                                    {
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                //check email sent out or not 
                                if (this.PlcBeatMailSent == "Y")
                                {
                                    ls_msg = "IESO PLC HeartBeat restored at: " + getStndSummerDateTime().ToString() +
                                                ". PlcCounterRd: " + resultList[0].ToString() +
                                                "; PlcCounterWr: " + resultList[1].ToString();

                                    actionLog("DspProcess:DspPLCHeartBeat:", ls_msg);

                                    ls_emailSubject = "PLC HeartBeat Active has been restored";
                                    ls_emailBody = ls_msg;

                                    li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                                    if (li_emailSent < 0)
                                    {
                                        li_return = -1;
                                        break;
                                    }
                                    else
                                    {
                                        this.PlcBeatMailSent = "N";
                                    }
                                }
                                else
                                {
                                    ls_msg = "IESO PLC HeartBeat Active at: " + getStndSummerDateTime().ToString() +
                                                ". PlcCounterRd: " + resultList[0].ToString() +
                                                "; PlcCounterWr: " + resultList[1].ToString();

                                    actionLog("DspProcess:DspPLCHeartBeat:", ls_msg);
                                }
                            }
                        }
                        else
                        {
                            ls_msg = "IESO PLC HeartBeat Active at: " + getStndSummerDateTime().ToString() +
                                        ". PlcCounterRd: " + resultList[0].ToString() +
                                        "; PlcCounterWr: " + resultList[1].ToString();
                            actionLog("DspProcess:DspPLCHeartBeat:", ls_msg);
                        }

                        this.PlcCntWrPrev = li_plccounterwr;

                        //update PLC sendtoPLC counter value
                        parmName = null;
                        parmType = null;
                        parmValue = null;
                        parmSize = null;

                        parmName = new ArrayList();
                        parmType = new ArrayList();
                        parmValue = new ArrayList();
                        parmSize = new ArrayList();

                        //continue to update PLC Rd Counter
                        li_plccounterrd += 1;

                        if (li_plccounterrd > 500)
                        {
                            li_plccounterrd = 1;
                        }

                        parmName.Add("@dispatch_on");
                        parmType.Add(OleDbType.Char);
                        parmValue.Add("PlcCounterRd");
                        parmSize.Add(16);

                        parmName.Add("@dispatch_value");
                        parmType.Add(OleDbType.Integer);

                        parmValue.Add(li_plccounterrd);
                        parmSize.Add(10);

                        if (oleCnn.State == ConnectionState.Closed)
                        {
                            oleCnn.Open();
                        }

                        actionLog("DspProcess:DspPLCHeartBeat:", "Executing ssp_update_plc_dsp_count.");
                        parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                        li_return = dbcomm.updOleProData("ssp_update_plc_dsp_count", parmList, oleCnn);
                        actionLog("DspProcess:DspPLCHeartBeat:", "Done with ssp_update_plc_dsp_count (" + dt.Rows.Count.ToString() + ")");
                    }
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspPLCHeartBeat(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        private int DspLoadStatus()
        {
            string ls_msg = "";
            string ls_emailSubject = "";
            string ls_emailBody = "";
            int li_return = 1;
            int li_emailSent = 0;
            int li_cntdown_min = 0;
            int li_cntdown_sec = 0;
            int li_cntup_min = 0;
            int li_cntup_sec = 0;
            int li_dispatch_off = 0;
            int li_dispatch_on = 0;
            int li_eststartmin = 0;
            int li_or10mins = 0;
            int li_or30mins = 0;
            int li_oractivation = 0;
            int li_popup_msg = 0;
            int li_approval = 0;
            int li_trafficlight = 0;
            int li_dispatchfrom = 0;
            int li_emergency = 0;
            int li_power_on = 0;
            int li_overrideact = 0;
            int li_autodial1 = 0;
            int li_autodial2 = 0;
            int li_autodial3 = 0;
            int li_autodial4 = 0;
            int li_autodialfail = 0;
            int li_popupmsg1 = 0;
            int li_2000active = 0;
            int li_eststartpopup = 0;
            int li_dspcontroloff = 0;
            int li_dspontimeout = 0;
            int li_dsponestmin = 0;
            int li_dsponexpire = 0;
            int li_startpointday = 0;
            int li_startpointhr = 0;
            int li_startpointmonth = 0;
            int li_plccounter = 0;
            int li_plccounterrd = 0;
            int li_plccounterwr = 0;

            ArrayList parmList = new ArrayList();
            ArrayList parmName = new ArrayList();
            ArrayList parmType = new ArrayList();
            ArrayList parmValue = new ArrayList();
            ArrayList parmSize = new ArrayList();
            ArrayList resultList = new ArrayList();

            DataTable dt = new DataTable();

            try
            {
                //1.get Dsp Load status from target database
                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }

                actionLog("DspProcess:DspLoadStatus:", "Executing ssp_get_dispatch_onoff .");
                dt = dbcomm.getOleProData("ssp_get_dispatch_onoff ", oleCnn);
                actionLog("DspProcess:DspLoadStatus:", "Done with ssp_get_dispatch_onoff (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        foreach (DataColumn myCol in dt.Columns)
                        {
                            resultList.Add(myRow[myCol]);
                        }

                        actionLog("DspProcess:DspLoadStatus:", "DspLoadStatus values---------------------------------------------------");
                        actionLog("DspProcess:DspLoadStatus:", "cntdown_min = " + resultList[0].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "cntdown_sec = " + resultList[1].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "cntup_min = " + resultList[2].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "cntup_sec = " + resultList[3].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "dispatch_off = " + resultList[4].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "dispatch_on = " + resultList[5].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "eststartmin = " + resultList[6].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "or10mins = " + resultList[7].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "or30mins = " + resultList[8].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "oractivation = " + resultList[9].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "popup_msg = " + resultList[10].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "approval = " + resultList[11].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "trafficlight = " + resultList[12].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "dispatchfrom = " + resultList[13].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "emergency = " + resultList[14].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "power_on = " + resultList[15].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "overrideact = " + resultList[16].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "autodial1 = " + resultList[17].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "autodial2 = " + resultList[18].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "autodial3 = " + resultList[19].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "autodial4 = " + resultList[20].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "autodialfail = " + resultList[21].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "popupmsg1 = " + resultList[22].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "2000active = " + resultList[23].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "eststartpopup = " + resultList[24].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "dspcontroloff = " + resultList[25].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "dspontimeout = " + resultList[26].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "dsponestmin = " + resultList[27].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "dsponexpire = " + resultList[28].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "startpointday = " + resultList[29].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "startpointhr = " + resultList[30].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "startpointmonth = " + resultList[31].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "plccounter = " + resultList[32].ToString());
                        actionLog("DspProcess:DspLoadStatus:", "DspCountDownMinPrev = " + this.DspCountDownMinPrev.ToString());
                        actionLog("DspProcess:DspLoadStatus:", "DspCountDownSecPrev = " + this.DspCountDownSecPrev.ToString());
                        actionLog("DspProcess:DspLoadStatus:", "END--------------------------------------------------------------------");

                        li_cntdown_min = (int)resultList[0];
                        li_cntdown_sec = (int)resultList[1];
                        li_cntup_min = (int)resultList[2];
                        li_cntup_sec = (int)resultList[3];
                        li_dispatch_off = (int)resultList[4];
                        li_dispatch_on = (int)resultList[5];
                        li_eststartmin = (int)resultList[6];
                        li_or10mins = (int)resultList[7];
                        li_or30mins = (int)resultList[8];
                        li_oractivation = (int)resultList[9];
                        li_popup_msg = (int)resultList[10];
                        li_approval = (int)resultList[11];
                        li_trafficlight = (int)resultList[12];
                        li_dispatchfrom = (int)resultList[13];
                        li_emergency = (int)resultList[14];
                        li_power_on = (int)resultList[15];
                        li_overrideact = (int)resultList[16];
                        li_autodial1 = (int)resultList[17];
                        li_autodial2 = (int)resultList[18];
                        li_autodial3 = (int)resultList[19];
                        li_autodial4 = (int)resultList[20];
                        li_autodialfail = (int)resultList[21];
                        li_popupmsg1 = (int)resultList[22];
                        li_2000active = (int)resultList[23];
                        li_eststartpopup = (int)resultList[24];
                        li_dspcontroloff = (int)resultList[25];
                        li_dspontimeout = (int)resultList[26];
                        li_dsponestmin = (int)resultList[27];
                        li_dsponexpire = (int)resultList[28];
                        li_startpointday = (int)resultList[29];
                        li_startpointhr = (int)resultList[30];
                        li_startpointmonth = (int)resultList[31];
                        li_plccounter = (int)resultList[32];

                        if (li_eststartmin == 0)
                        {
                            this.DspEstMinCount = 0;
                        }

                        if (li_dspcontroloff == 0)
                        {
                            if (this.CntlOffMailSent == "N")
                            {
                                ls_msg = "DLS - Over-Ride Key Switch Activated at: " + getStndSummerDateTime().ToString();

                                actionLog("DspProcess:DspLoadStatus:", ls_msg);

                                ls_emailSubject = "DSPCNTLOFF";
                                ls_emailBody = ls_msg;
                                li_emailSent = sendDbEmail("DSPCNTLOFF", "DSPCNTLOFF", ls_emailSubject, ls_emailBody);

                                if (li_emailSent < 0)
                                {
                                    li_return = -1;
                                    break;
                                }
                                else
                                {
                                    //assign mailsent indicator
                                }

                                //1.1 update target database action log table
                                li_return = DspDbActionLog("DSPCNTLOFF", ls_emailBody);
                                if (li_return < 0)
                                {
                                    break;
                                }

                                this.CntlOffMailSent = "Y";
                            }
                        }
                        else
                        {
                            if (this.CntlOffMailSent == "Y")
                            {
                                ls_msg = "DLS - Over-Ride Key Switch deactivated at: " + getStndSummerDateTime().ToString();

                                actionLog("DspProcess:DspLoadStatus:", ls_msg);

                                ls_emailSubject = "DSPCNTLOFF";
                                ls_emailBody = ls_msg;
                                li_emailSent = sendDbEmail("DSPCNTLOFF", "DSPCNTLOFF", ls_emailSubject, ls_emailBody);

                                if (li_emailSent < 0)
                                {
                                    li_return = -1;
                                    break;
                                }
                                else
                                {
                                    //assign meailsent indicator
                                }

                                //update target database action log table
                                li_return = DspDbActionLog("DSPCNTLOFF", ls_emailBody);

                                if (li_return < 0)
                                {
                                    break;
                                }
                                this.CntlOffMailSent = "N";
                            }
                        }

                        li_plccounterrd = Int32.Parse(resultList[0].ToString());
                        li_plccounterwr = Int32.Parse(resultList[1].ToString());

                        //2. check estimate start minutes 0 value under countdown minutes +countdown second >0
                        if (li_2000active == 0 & li_approval == 0 & li_dispatchfrom == 1 & li_dispatch_off == 1 & li_dspcontroloff == 1 & li_emergency == 0 & li_power_on == 0)
                        {
                            if (((li_cntdown_min * 60 + li_cntdown_sec) > 0) & ((li_cntdown_min * 60 + li_cntdown_sec) < (this.DspCountDownMinPrev * 60 + this.DspCountDownSecPrev)))
                            {
                                int li_test = (li_cntdown_min * 60 + li_cntdown_sec);
                                int li_test1 = (this.DspCountDownMinPrev * 60 + this.DspCountDownSecPrev);

                                actionLog("DspProcess:DspLoadStatus:", "Count Down Seconds: " + li_test.ToString() +
                                                "; Previous Count Down Seconds: " + li_test1.ToString() +
                                                "; Estimate Start Minutes: " + li_eststartmin.ToString());

                                if (li_eststartmin > 0)
                                {
                                    //check if Est Min Count >5.
                                    //we give system 50 seconds to process Estminutes value
                                    if (this.DspEstMinCount > 5)
                                    {
                                        parmName = null;
                                        parmType = null;
                                        parmValue = null;
                                        parmSize = null;

                                        parmName = new ArrayList();
                                        parmType = new ArrayList();
                                        parmValue = new ArrayList();
                                        parmSize = new ArrayList();

                                        //update EstStartMins to force to 0

                                        parmName.Add("@dispatch_on");
                                        parmType.Add(OleDbType.Char);
                                        parmValue.Add("EstStartMins");
                                        parmSize.Add(16);

                                        parmName.Add("@dispatch_value");
                                        parmType.Add(OleDbType.Integer);
                                        parmValue.Add(0);
                                        parmSize.Add(10);

                                        if (oleCnn.State == ConnectionState.Closed)
                                        {
                                            oleCnn.Open();
                                        }

                                        actionLog("DspProcess:DspLoadStatus:", "Executing ssp_update_plc_dsp_dispatch to force EstStartMins = 0");
                                        parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                                        li_return = dbcomm.updOleProData("ssp_update_plc_dsp_dispatch", parmList, oleCnn);
                                        actionLog("DspProcess:DspLoadStatus:", "Done with ssp_update_plc_dsp_dispatch (" + li_return.ToString() + ")");

                                        this.DspEstMinCount = 0;
                                    }
                                    else
                                    {
                                        this.DspEstMinCount = this.DspEstMinCount + 1;
                                    }
                                }
                            }
                        }

                        //3. check dispatchfrom parameter. if it equal 1 then update to 0 under approval = 1
                        if (li_approval == 1 & li_dispatchfrom == 1)
                        {
                            if (this.DispatchFromCount > 3)
                            {
                                parmName = null;
                                parmType = null;
                                parmValue = null;
                                parmSize = null;

                                parmName = new ArrayList();
                                parmType = new ArrayList();
                                parmValue = new ArrayList();
                                parmSize = new ArrayList();

                                parmName.Add("@dispatch_on");
                                parmType.Add(OleDbType.Char);
                                parmValue.Add("DispatchFrom");
                                parmSize.Add(16);

                                parmName.Add("@dispatch_value");
                                parmType.Add(OleDbType.Integer);
                                parmValue.Add(0);
                                parmSize.Add(10);

                                if (oleCnn.State == ConnectionState.Closed)
                                {
                                    oleCnn.Open();
                                }

                                actionLog("DspProcess:DspLoadStatus:", "Executing ssp_update_plc_dsp_dispatch to force EstStartMins = 0");
                                parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                                li_return = dbcomm.updOleProData("ssp_update_plc_dsp_dispatch", parmList, oleCnn);
                                actionLog("DspProcess:DspLoadStatus:", "Done with ssp_update_plc_dsp_dispatch (" + li_return.ToString() + ")");

                                this.DispatchFromCount = 0;
                            }
                            else
                            {
                                this.DispatchFromCount += 1;
                            }
                        }

                        //4. update 20 minutes to due act status
                        if (li_2000active == 0 & li_approval == 0 & li_cntdown_min == 20 & li_dispatchfrom == 1 & li_dispatch_off == 1 & li_dspcontroloff == 1 & li_emergency == 0 & li_power_on == 0)
                        {
                            if (this.Minute20Remind == "N")
                            {
                                parmName = null;
                                parmType = null;
                                parmValue = null;
                                parmSize = null;

                                parmName = new ArrayList();
                                parmType = new ArrayList();
                                parmValue = new ArrayList();
                                parmSize = new ArrayList();

                                //continue to update PLC Rd Counter
                                li_plccounterrd = li_plccounterrd + 1;
                                if (li_plccounterrd > 500)
                                {
                                    li_plccounterrd = 1;
                                }

                                parmName.Add("@dispatch_on");
                                parmType.Add(OleDbType.Char);
                                parmValue.Add("PopUpMsg");
                                parmSize.Add(16);

                                parmName.Add("@dispatch_value");
                                parmType.Add(OleDbType.Integer);
                                parmValue.Add(0);
                                parmSize.Add(10);

                                if (oleCnn.State == ConnectionState.Closed)
                                {
                                    oleCnn.Open();
                                }

                                actionLog("DspProcess:DspLoadStatus:", "Executing ssp_update_plc_dsp_dispatch.");
                                parmList = dbcomm.MakeInParm(parmName, parmType, parmSize, parmValue);
                                li_return = dbcomm.updOleProData("ssp_update_plc_dsp_dispatch", parmList, oleCnn);
                                actionLog("DspProcess:DspLoadStatus:", "Done with ssp_update_plc_dsp_dispatch (" + li_return.ToString() + "(");

                                this.Minute20Remind = "Y";
                            }
                        }
                        else
                        {
                            this.Minute20Remind = "N";
                        }

                        //check count down and cout up.
                        li_return = DspFlinkCount(resultList);

                        this.DspCountDownMinPrev = li_cntdown_min;
                        this.DspCountDownSecPrev = li_cntdown_sec;
                    }
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspLoadStatus(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return li_return;
        }
        public ArrayList ParmPassBack()
        {
            ArrayList parmList = new ArrayList();
            //
            // TODO: Assign parameters to the references.
            //

            actionLog("DspProcess:ParmPassBack:", "The DspProcess out parameters------------------------------------------");
            actionLog("DspProcess:ParmPassBack:", "PlcCntWrDatePrev: " + this.PlcCntWrDatePrev.ToString());
            actionLog("DspProcess:ParmPassBack:", "PlcCntWrPrev: " + this.PlcCntWrPrev.ToString());
            actionLog("DspProcess:ParmPassBack:", "PlcBeatMailSent: " + this.PlcBeatMailSent.ToString());
            actionLog("DspProcess:ParmPassBack:", "IesoBeatMailSent: " + this.IesoBeatMailSent.ToString());
            actionLog("DspProcess:ParmPassBack:", "Minute20Remind: " + this.Minute20Remind.ToString());
            actionLog("DspProcess:ParmPassBack:", "CntlOffMailSent: " + this.CntlOffMailSent.ToString());
            actionLog("DspProcess:ParmPassBack:", "IESOUnixTime: " + this.IESOUnixTime.ToString());
            actionLog("DspProcess:ParmPassBack:", "IESOUnixDateTime: " + this.IESOUnixDateTime.ToString());
            actionLog("DspProcess:ParmPassBack:", "IesoORMarket: " + this.IesoOrMarket.ToString());
            actionLog("DspProcess:ParmPassBack:", "PlcCountUp: " + this.PlcCountUp.ToString());
            actionLog("DspProcess:ParmPassBack:", "PlcCountDown: " + this.PlcCountDown.ToString());
            actionLog("DspProcess:ParmPassBack:", "PlcCntUpDatePrev: " + this.PlcCntUpDatePrev.ToString());
            actionLog("DspProcess:ParmPassBack:", "PlcCntDownDatePrev: " + this.PlcCntDownDatePrev.ToString());
            actionLog("DspProcess:ParmPassBack:", "FlinkCntUpMailSent: " + this.FlinkCntUpMailSent.ToString());
            actionLog("DspProcess:ParmPassBack:", "FlinkCntDownMailSent: " + this.FlinkCntDownMailSent.ToString());
            actionLog("DspProcess:ParmPassBack:", "PwrOnBeatMailSent: " + this.PwrOnBeatMailSent.ToString());
            actionLog("DspProcess:ParmPassBack:", "DspCountDownMinPrev: " + this.DspCountDownMinPrev.ToString());
            actionLog("DspProcess:ParmPassBack:", "DspCountDownSecPrev: " + this.DspCountDownSecPrev.ToString());
            actionLog("DspProcess:ParmPassBack:", "DspEstMinCount: " + this.DspEstMinCount.ToString());
            actionLog("DspProcess:ParmPassBack:", "dispatchFromCount: " + this.DispatchFromCount.ToString());
            actionLog("DspProcess:ParmPassBack:", "End--------------------------------------------------------------------");

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

            return parmList;
        }
        public DateTime GetServerTime()
        {
            DataTable dt = new DataTable();
            DateTime ldt_date = DateTime.Parse("1990-01-01");

            try
            {
                if (oleCnn.State == ConnectionState.Closed)
                {
                    oleCnn.Open();
                }

                actionLog("DspProcess:GetServerTime:", "Executing ssp_get_server_time .");
                dt = dbcomm.getOleProData("ssp_get_server_time ", oleCnn);
                actionLog("DspProcess:GetServerTime:", "Done with ssp_get_server_time (" + dt.Rows.Count.ToString() + ")");

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow myRow in dt.Rows)
                    {
                        foreach (DataColumn myCol in dt.Columns)
                        {
                            ldt_date = (DateTime)myRow[myCol];
                            actionLog("DspProcess:GetServerTime:", "Database Server Time:" + ldt_date.ToString());
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:GetServerTime(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            finally
            {
                if (dt != null)
                {
                    dt.Dispose();
                }

                if (oleCnn.State == ConnectionState.Open)
                {
                    oleCnn.Close();
                }

                if (oleSrcCnn.State == System.Data.ConnectionState.Open)
                {
                    oleSrcCnn.Close();
                }
            }

            return ldt_date;
        }
        private int DspFlinkCount(System.Collections.ArrayList resultList)
        {
            string ls_msg = "";
            string ls_emailSubject = "";
            string ls_emailBody = "";
            int li_emailSent = 0;
            int li_return = 1;
            int li_cntdown_min = 0;
            int li_cntdown_sec = 0;
            int li_cntup_min = 0;
            int li_cntup_sec = 0;
            int li_dispatch_off = 0;
            int li_dispatch_on = 0;
            int li_eststartmin = 0;
            int li_or10mins = 0;
            int li_or30mins = 0;
            int li_oractivation = 0;
            int li_popup_msg = 0;
            int li_approval = 0;
            int li_trafficlight = 0;
            int li_dispatchfrom = 0;
            int li_emergency = 0;
            int li_power_on = 0;
            int li_overrideact = 0;
            int li_autodial1 = 0;
            int li_autodial2 = 0;
            int li_autodial3 = 0;
            int li_autodial4 = 0;
            int li_autodialfail = 0;
            int li_popupmsg1 = 0;
            int li_2000active = 0;
            int li_eststartpopup = 0;
            int li_dspcontroloff = 0;
            int li_dspontimeout = 0;
            int li_dsponestmin = 0;
            int li_dsponexpire = 0;
            int li_startpointday = 0;
            int li_startpointhr = 0;
            int li_startpointmonth = 0;
            int li_plccounter = 0;
            int li_flinkcntdown = 0;
            int li_flinkcntup = 0;
            int li_timespan_curr = 0;

            try
            {
                li_cntdown_min = (int)resultList[0];
                li_cntdown_sec = (int)resultList[1];
                li_cntup_min = (int)resultList[2];
                li_cntup_sec = (int)resultList[3];
                li_dispatch_off = (int)resultList[4];
                li_dispatch_on = (int)resultList[5];
                li_eststartmin = (int)resultList[6];
                li_or10mins = (int)resultList[7];
                li_or30mins = (int)resultList[8];
                li_oractivation = (int)resultList[9];
                li_popup_msg = (int)resultList[10];
                li_approval = (int)resultList[11];
                li_trafficlight = (int)resultList[12];
                li_dispatchfrom = (int)resultList[13];
                li_emergency = (int)resultList[14];
                li_power_on = (int)resultList[15];
                li_overrideact = (int)resultList[16];
                li_autodial1 = (int)resultList[17];
                li_autodial2 = (int)resultList[18];
                li_autodial3 = (int)resultList[19];
                li_autodial4 = (int)resultList[20];
                li_autodialfail = (int)resultList[21];
                li_popupmsg1 = (int)resultList[22];
                li_2000active = (int)resultList[23];
                li_eststartpopup = (int)resultList[24];
                li_dspcontroloff = (int)resultList[25];
                li_dspontimeout = (int)resultList[26];
                li_dsponestmin = (int)resultList[27];
                li_dsponexpire = (int)resultList[28];
                li_startpointday = (int)resultList[29];
                li_startpointhr = (int)resultList[30];
                li_startpointmonth = (int)resultList[31];
                li_plccounter = (int)resultList[32];

                if ((li_dspcontroloff == 1) && (li_2000active == 0) && (li_approval == 1) && (li_overrideact == 0) && (li_emergency == 0))
                {
                    #region Check Dispatch OFF and Power ON

                    //dispatch off and power on
                    if ((li_dispatch_off == 1) && (li_power_on == 1))
                    {
                        //there is nothing to do with count up indicators we have to check countupsendemail indicator first

                        //check email sent out or not 
                        if (this.FlinkCntUpMailSent == "Y")
                        {
                            ls_msg = "PLC Dispatch On Count Up restored at: " + getStndSummerDateTime().ToString() +
                                        ". CountUpMin: " + li_cntup_min.ToString() +
                                        "; CountUpSec: " + li_cntup_sec.ToString();

                            actionLog("DspProcess:DspFlinkCount:", ls_msg);

                            ls_emailSubject = "PLC Dispatch On Count Up has been restored";
                            ls_emailBody = ls_msg;
                            li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                            if (li_emailSent < 0)
                            {
                                li_return = -1;
                            }
                            else
                            {
                                this.FlinkCntUpMailSent = "N";
                            }
                        }

                        this.FlinkCntUpMailSent = "N";
                        this.PlcCntUpDatePrev = this.getStndSummerDateTime();

                        li_flinkcntdown = li_cntdown_min * 60 + li_cntdown_sec;

                        if (this.PlcCountDown != li_flinkcntdown)
                        {
                            actionLog("DspProcess:DspFlinkCount:", "Count Down value in seconds: " + li_flinkcntdown.ToString() +
                                            "; CurDate: " + getStndSummerDateTime().ToString() +
                                            "; Previous Count Down Date: " + this.PlcCntDownDatePrev.ToString());

                            this.PlcCntDownDatePrev = getStndSummerDateTime();
                        }

                        if (li_flinkcntdown >= 0)
                        {
                            if (this.PlcCountDown == li_flinkcntdown)
                            {
                                TimeSpan span = (getStndSummerDateTime()).Subtract(this.PlcCntDownDatePrev);

                                li_timespan_curr = (int)span.TotalMinutes;

                                actionLog("DspProcess:DspFlinkCount:", "Count Down time span in minutes: " + li_timespan_curr.ToString());

                                int li_beat_timeout = Int32.Parse(Panel.strPlcHeartBeatMinute);

                                //check time out
                                if (li_timespan_curr >= li_beat_timeout)
                                {
                                    if (this.FlinkCntDownMailSent == "N")
                                    {
                                        ls_msg = "PLC Dispatch Off Count Down Not Active at: " + this.PlcCntDownDatePrev.ToString() +
                                                    ". CountDownMin: " + li_cntdown_min.ToString() +
                                                    "; CountDownSec: " + li_cntdown_sec.ToString();

                                        actionLog("DspProcess:DspFlinkCount:", ls_msg);

                                        ls_emailSubject = "PLC Count Down Not Active";
                                        ls_emailBody = ls_msg;
                                        li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                                        if (li_emailSent < 0)
                                        {
                                            li_return = -1;
                                        }
                                        else
                                        {
                                            this.FlinkCntDownMailSent = "Y";
                                        }

                                        li_return = DspDbActionLog("PLCHEARTBEAT", ls_emailBody);
                                        if (li_return < 0)
                                        {
                                            li_return = -1;
                                        }
                                    }
                                }
                                else
                                {
                                    if (this.FlinkCntDownMailSent == "Y")
                                    {
                                        ls_msg = "PLC Dispatch Off Count Down restored at: " + getStndSummerDateTime().ToString() +
                                                    ". CountDownMin: " + li_cntdown_min.ToString() +
                                                    "; CountDownSec: " + li_cntdown_sec.ToString();

                                        actionLog("DspProcess:DspFlinkCount:", ls_msg);

                                        ls_emailSubject = "PLC Dispatch Off Count Down has been restored";
                                        ls_emailBody = ls_msg;
                                        li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                                        if (li_emailSent < 0)
                                        {
                                            li_return = -1;
                                        }
                                        else
                                        {
                                            this.FlinkCntDownMailSent = "N";
                                        }
                                    }
                                    else
                                    {
                                        ls_msg = "PLC Dispatch Off Count Down Active at: " + getStndSummerDateTime().ToString() +
                                                    ". CountDownMin: " + li_cntdown_min.ToString() +
                                                    "; CountDownSec: " + li_cntdown_sec.ToString();

                                        actionLog("DspProcess:DspFlinkCount:", ls_msg);
                                    }
                                }
                            }
                            else
                            {
                                ls_msg = "PLC Dispatch Off Count Down Active at: " + getStndSummerDateTime().ToString() +
                                            ". CountDownMin: " + li_cntdown_min.ToString() +
                                            "; CountDownSec: " + li_cntdown_sec.ToString();

                                actionLog("DspProcess:DspFlinkCount:", ls_msg);
                            }
                        }
                    }

                    if ((li_dispatch_off == 1) && (li_power_on == 0))
                    {
                        if (this.FlinkCntDownMailSent == "Y")
                        {
                            ls_msg = "PLC Dispatch Off Count Down restored at: " + getStndSummerDateTime().ToString();

                            actionLog("DspProcess:DspFlinkCount:", ls_msg);

                            ls_emailSubject = "PLC Dispatch Off Count Down has been restored";
                            ls_emailBody = ls_msg;
                            li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                            if (li_emailSent < 0)
                            {
                                li_return = -1;
                            }
                            else
                            {
                                this.FlinkCntDownMailSent = "N";
                            }
                        }

                        //dispatch off and power on but the count down = 0
                        this.FlinkCntDownMailSent = "N";
                        this.PlcCntDownDatePrev = this.getStndSummerDateTime();
                        this.FlinkCntUpMailSent = "N";
                        this.PlcCntUpDatePrev = this.getStndSummerDateTime();
                    }

                    this.PlcCountDown = li_flinkcntdown;

                    #endregion

                    #region Dispatch ON and Power OFF

                    if ((li_dispatch_on == 1) && (li_power_on == 0))
                    {
                        //there is nothing to do with coutdown indicators
                        //we have to check countdown sentemail indicator first

                        if (this.FlinkCntDownMailSent == "Y")
                        {
                            ls_msg = "PLC Dispatch Off Count Down restored at: " + getStndSummerDateTime().ToString();

                            actionLog("DspProcess:DspFlinkCount:", ls_msg);

                            ls_emailSubject = "PLC Dispatch Off Count Down has been restored";
                            ls_emailBody = ls_msg;
                            li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                            if (li_emailSent < 0)
                            {
                                li_return = -1;
                            }
                            else
                            {
                                this.FlinkCntDownMailSent = "N";
                            }
                        }

                        //set count down indicators to default
                        this.FlinkCntDownMailSent = "N";
                        this.PlcCntDownDatePrev = this.getStndSummerDateTime();

                        li_flinkcntup = li_cntup_min * 60 + li_cntup_sec;

                        if (this.PlcCountUp != li_flinkcntup)
                        {
                            actionLog("DspProcess:DspFlinkCount:", "PLC Count Up values seconds: " + li_flinkcntup.ToString() +
                                            "; CurDate: " + getStndSummerDateTime().ToString() +
                                            "; FlinkCntUpDatePrev: " + this.PlcCntUpDatePrev.ToString());

                            this.PlcCntUpDatePrev = getStndSummerDateTime();
                        }

                        if (li_flinkcntup >= 0)
                        {
                            if (this.PlcCountUp == li_flinkcntup)
                            {
                                TimeSpan span = (getStndSummerDateTime()).Subtract(this.PlcCntUpDatePrev);

                                li_timespan_curr = (int)span.TotalMinutes;

                                actionLog("DspProcess:DspFlinkCount:", "PLC Count Up time span minutes: " + li_timespan_curr.ToString());

                                int li_beat_timeout = Int32.Parse(Panel.strPlcHeartBeatMinute);

                                if (li_timespan_curr >= li_beat_timeout)
                                {
                                    if (this.FlinkCntUpMailSent == "N")
                                    {
                                        ls_msg = "PLC Dispatch On Count Up Not Active at: " + this.PlcCntUpDatePrev.ToString() +
                                                    ". CountUpMin: " + li_cntup_min.ToString() +
                                                    "; CountUpSec: " + li_cntup_sec.ToString();

                                        actionLog("DspProcess:DspFlinkCount:", ls_msg);

                                        ls_emailSubject = "PLC Count Up Not Active";
                                        ls_emailBody = ls_msg;
                                        li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                                        if (li_emailSent < 0)
                                        {
                                            li_return = -1;
                                        }
                                        else
                                        {
                                            this.FlinkCntUpMailSent = "Y";
                                        }

                                        li_return = DspDbActionLog("PLCHEARTBEAT", ls_emailBody);
                                        if (li_return < 0)
                                        {
                                            li_return = -1;
                                        }
                                    }
                                }
                                else
                                {
                                    if (this.FlinkCntUpMailSent == "Y")
                                    {
                                        ls_msg = "PLC Dispatch On Count Up restored at: " + getStndSummerDateTime().ToString() +
                                                    ". CountDownMin: " + li_cntup_min.ToString() +
                                                    "; CountDownSec: " + li_cntup_sec.ToString();

                                        actionLog("DspProcess:DspFlinkCount:", ls_msg);

                                        ls_emailSubject = "PLC Dispatch On Count Up has been restored";
                                        ls_emailBody = ls_msg;
                                        li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                                        if (li_emailSent < 0)
                                        {
                                            li_return = -1;
                                        }
                                        else
                                        {
                                            this.FlinkCntUpMailSent = "N";
                                        }
                                    }
                                    else
                                    {
                                        ls_msg = "PLC Dispatch On Count Up Active at: " + getStndSummerDateTime().ToString() +
                                                    ". CountUpMin: " + li_cntup_min.ToString() +
                                                    "; CountUpSec: " + li_cntup_sec.ToString();

                                        actionLog("DspProcess:DspFlinkCount:", ls_msg);
                                    }
                                }
                            }
                            else
                            {
                                ls_msg = "PLC Dispatch On Count Up Active at: " + getStndSummerDateTime().ToString() +
                                            ". CountUpMin: " + li_cntup_min.ToString() +
                                            "; CountUpSec: " + li_cntup_sec.ToString();

                                actionLog("DspProcess:DspFlinkCount:", ls_msg);
                            }
                        }
                    }

                    if ((li_dispatch_on == 1) && (li_power_on == 1))
                    {
                        if (this.FlinkCntUpMailSent == "Y")
                        {
                            ls_msg = "PLC Dispatch On Count Up restored at: " + getStndSummerDateTime().ToString();

                            actionLog("DspProcess:DspFlinkCount:", ls_msg);

                            ls_emailSubject = "PLC Dispatch On Count Up has been restored";
                            ls_emailBody = ls_msg;
                            li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                            if (li_emailSent < 0)
                            {
                                li_return = -1;
                            }
                            else
                            {
                                this.FlinkCntUpMailSent = "N";
                            }
                        }

                        //dispatch on and power off but the count down =0
                        this.FlinkCntUpMailSent = "N";
                        this.PlcCntUpDatePrev = this.getStndSummerDateTime();
                        this.FlinkCntDownMailSent = "N";
                        this.PlcCntDownDatePrev = this.getStndSummerDateTime();
                    }

                    //after done checking Flink Count Down, load Flink Count Down value to the Object
                    this.PlcCountUp = li_flinkcntup;

                    #endregion
                }
                else
                {
                    if (this.FlinkCntUpMailSent == "Y")
                    {
                        ls_msg = "PLC Dispatch On Count Up restored at: " + getStndSummerDateTime().ToString() +
                                    ". CountUpMin: " + li_cntup_min.ToString() +
                                    "; CountUpSec: " + li_cntup_sec.ToString();

                        actionLog("DspProcess:DspFlinkCount:", ls_msg);

                        ls_emailSubject = "PLC Dispatch On Count Up has been restored";
                        ls_emailBody = ls_msg;
                        li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                        if (li_emailSent < 0)
                        {
                            li_return = -1;
                        }
                        else
                        {
                            this.FlinkCntUpMailSent = "N";
                        }
                    }

                    if (this.FlinkCntDownMailSent == "Y")
                    {
                        ls_msg = "PLC Dispatch Off Count Down restored at: " + getStndSummerDateTime().ToString() +
                                    ". CountDownMin: " + li_cntdown_min.ToString() +
                                    "; CountDownSec: " + li_cntdown_sec.ToString();

                        actionLog("DspProcess:DspFlinkCount:", ls_msg);

                        ls_emailSubject = "PLC Dispatch Off Count Down has been restored";
                        ls_emailBody = ls_msg;
                        li_emailSent = sendDbEmail("PLCHEARTBEAT", "PLCHEARTBEAT", ls_emailSubject, ls_emailBody);

                        if (li_emailSent < 0)
                        {
                            li_return = -1;
                        }
                        else
                        {
                            this.FlinkCntDownMailSent = "N";
                        }
                    }

                    this.FlinkCntUpMailSent = "N";
                    this.PlcCntUpDatePrev = this.getStndSummerDateTime();
                    this.FlinkCntDownMailSent = "N";
                    this.PlcCntDownDatePrev = this.getStndSummerDateTime();
                }
            }

            catch (Exception ex)
            {
                li_return = -1;

                StackTrace st = new StackTrace(new StackFrame(true));
                this.actionLog("DspProcess:DspFlinkCount(" + st.GetFrame(0).GetFileLineNumber().ToString() + "):", "Exception Error: " + ex.ToString());
            }

            return li_return;
        }
    }
}
