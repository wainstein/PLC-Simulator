using PLCTools.Common;
using PLCTools.Components;
using PLCTools.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Timer = System.Timers.Timer;

namespace PLCTools
{
    public partial class Panel : Form
    {
        delegate void SetStatusLabelCallback(string text);
        delegate void SetTextCallback(Control control, string text);
        delegate void SetEnableCallback(Control control, bool state);
        public static Queue<string> dspBatchLog { get; set; } = new Queue<string>();
        public static Queue<string> dspBatchMessage { get; set; } = new Queue<string>();
        public static Queue<string> dspBatchMail { get; set; } = new Queue<string>();
        private bool _isStarting { get; set; } = false;
        private bool _isConnecting { get; set; } = false;
        private DspBatch batch = new DspBatch();
        private PLCAutomation plcAuto;
        private static Timer timerPLC { get; set; } = new Timer();
        private static int inTimer = 0;
        public static string strActionLogPath { get; set; }
        public static string strActionLogDeleteDays { get; set; }
        public static string strConnectionSource { get; set; }
        public static string strConnectionTarget { get; set; }
        public static string strMailHost { get; set; }
        public static string strTimerSecond { get; set; }
        public static string strServerIP { get; set; }
        public static string strSendPort { get; set; }
        public static string strMailFrom { get; set; }
        public static string strMailLogin { get; set; }
        public static int strMailPort { get; set; }
        public static string strMailTo { get; set; }
        public static bool isSendMail { get; set; }
        public static bool isSSLMail { get; set; }
        public static string strPlcHeartBeatMinute { get; set; }
        public static string strWinMsgSendInd { get; set; }
        public static string strMailPass { get; set; }
        private static bool _opcState { get; set; } = false;
        private static bool _opcSubState { get; set; } = false;
        private static bool _dspBatchState { get; set; } = false;
        private static bool _dspBatchSubState { get; set; } = false;
        private static bool _q2DBState { get; set; } = false;
        private static bool _q2DBSubState { get; set; } = false;
        internal static Dictionary<string, string> messages { get; set; } = new Dictionary<string, string> {
            { "OPENCONN","reading from DB"},
            { "RETRIVE","Retriving data from database" },
            { "DBSUCCESS","Reading completed" },
            { "DBFAIL","DB connection failed" },
            { "CONNOPC","Connecting to OPC... " },
            { "READOPC","Reading from PLC..." },
            { "OPCREAD","Reading successful" },
            { "WRITEOPC","Writing to OPC... " },
            { "OPCWROTE", "Write Completed" },
            { "WAITEST","Need input the power on time... " },
            { "STB","Standby..." },
            { "DISCONN", "Disconnected"}
        };
        public static Dictionary<int, string> indicators { get; set; } = new Dictionary<int, string>
        {
            {-1, "DSP_AUTO_DIAL1"},
            {0, "DSP_AUTO_DIAL2" },
            {1, "DSP_AUTO_DIAL3"},
            {2, "DSP_AUTO_DIAL4" },
            {7, "DSP_2000_BID_ACTIVE"},
            {8, "DSP_OR_30MIN" },
            {9, "DSP_OR_10MIN" },
            {10, "DSP_DISPATCH_ON" },
            {11, "DSP_FORCE_DISP_OFF" },
            {12, "DSP_EMERG_OVERRIDE_ACT" },
            {13, "DSP_OR_ACT" },
            {14, "DSP_DISPATCH_OFF" }
        };
        public Panel()
        {
            InitializeComponent();
            foreach (SwitchControl item in this.PLCDashboard.Controls.OfType<SwitchControl>().Where(item => indicators.Values.Contains(item.Name)))
            {
                item.DisableSwitch();
            }
        }
        private BindingList<OPCItems> initializing()
        {
            BindingList<OPCItems> plcData = new BindingList<OPCItems>();
            setEnabled(this.Stop, false);
            setEnabled(this.panelStart, false);
            setEnabled(this.Start, false);
            {
                string sqlConnection = "Server='" + this.DBAddr.Text + "'; Database='OPC2DBMS'; User Id='" + this.DBUser.Text + "'; Password='" + this.DBPwd.Text + "'; MultipleActiveResultSets=true; Connect Timeout = 1";
                var sqlQuery = "select Tag, Register, Description, PLCName from PLC2DB_Tags where type = 'VAR' and OPCName = '" + OPCName.Text + "' and PLCName = '" + PLCName.Text + "'";
                try
                {
                    using (SqlConnection connection = new SqlConnection(sqlConnection))
                    {

                        setEnabled(this.Start, false);
                        plcAuto.PlcAutoLog.Enqueue(messages["RETRIVE"]);
                        connection.Open();
                        SqlCommand cmd = new SqlCommand(sqlQuery, connection);
                        SqlDataReader sdr = cmd.ExecuteReader();
                        while (sdr.Read())
                        {
                            OPCItems oPCItems = new OPCItems();
                            oPCItems.Tag = sdr[0].ToString().Trim();
                            oPCItems.Address = sdr[1].ToString().Trim();
                            oPCItems.Description = sdr[2].ToString().Trim();
                            oPCItems.PLCName = sdr[3].ToString().Trim();
                            plcData.Add(oPCItems);
                        }
                        sdr.Close();
                        toggleStartButtons(true);
                    }
                    plcAuto.PlcAutoLog.Enqueue(messages["DBSUCCESS"]);
                }
                catch (Exception ex)
                {
                    plcAuto.PlcAutoLog.Enqueue(messages["DBFAIL"]);
                    toggleStartButtons(false);
                    Console.WriteLine(ex);
                }
            }
            return plcData;
        }
        private void toggleStartButtons(bool state)
        {
            while (plcAuto.PlcAutoLog.Count > 0)
            {
                setStatusLabel(plcAuto.PlcAutoLog.Dequeue());
                Thread.Sleep(100);
            }
            _isStarting = state;
            setEnabled(this.Stop, state);
            setEnabled(this.tabControl, state);
            setEnabled(this.panelStart, !state);
            setEnabled(this.Start, !state);
        }
        private async void GlobalLogStart()
        {
            var t = Task.Run(() =>
             {
                 while (_isStarting || plcAuto.PlcAutoLog.Count > 0)
                 {
                     if (Quality.Text != plcAuto.OverallQuality) setText(Quality, plcAuto.OverallQuality);
                     if (plcAuto.PlcAutoLog.Count > 0)
                     {
                         setStatusLabel(plcAuto.PlcAutoLog.Dequeue());
                         Thread.Sleep(30);
                     }
                     else
                     {
                         Thread.Sleep(30);
                     }

                 }
             });
            await t;
            setStatusLabel("All services stopped.");
        }
        private async void refreshControls()
        {
            var t = Task.Run(() =>
            {
                bool commitNow = false;
                while (_isConnecting && _isStarting)
                {
                    foreach (SwitchControl item in this.PLCDashboard.Controls.OfType<SwitchControl>())
                    {
                        if (!item.fetch)
                        {
                            if (item.value != plcAuto.getTagValue(item.Name) || item.Pending)
                            {
                                item.value = plcAuto.getTagValue(item.Name);
                                item.quality = plcAuto.getTagItem(item.Name) == null ? 0 : (plcAuto.getTagItem(item.Name).Quality / 192.0);
                            }
                        }
                        else
                        {
                            plcAuto.queuePLCWrites(item.value, item.Name);
                            Thread.Sleep(200);
                            item.fetch = false;
                            commitNow = true;
                        }
                        if (item.Text != plcAuto.getTagValue(item.Name).ToString())
                        {
                            this.setText(item, plcAuto.getTagValue(item.Name).ToString());
                        }
                    }
                    foreach (TextBox item in this.PLCDashboard.Controls.OfType<TextBox>())
                    {
                        if (item.Text != plcAuto.getTagValue(item.Name).ToString())
                        {
                            this.setText(item, plcAuto.getTagValue(item.Name).ToString());
                        }
                    }

                    TimeSpan countdown = new TimeSpan(0);
                    TimeSpan countup = new TimeSpan(0);
                    if (plcAuto.isCountDown)
                    {
                        CountDown.BackColor = Color.RoyalBlue;
                        CountDown.ForeColor = Color.WhiteSmoke;
                        CountUp.BackColor = Color.WhiteSmoke;
                        CountUp.ForeColor = Color.Black;
                        countdown = plcAuto.countingDownTime - DateTime.Now;
                    }
                    else if (plcAuto.WaitEST)
                    {
                        CountDown.BackColor = Color.RoyalBlue;
                        CountDown.ForeColor = Color.WhiteSmoke;
                        CountUp.BackColor = Color.WhiteSmoke;
                        CountUp.ForeColor = Color.Black;
                    }
                    else if (plcAuto.isCountUp)
                    {
                        CountUp.BackColor = Color.RoyalBlue;
                        CountUp.ForeColor = Color.WhiteSmoke;
                        CountDown.BackColor = Color.WhiteSmoke;
                        CountDown.ForeColor = Color.Black;
                        countup = DateTime.Now - plcAuto.countingUpTime;
                    }
                    else
                    {
                        CountUp.BackColor = Color.WhiteSmoke;
                        CountUp.ForeColor = Color.Black;
                        CountDown.BackColor = Color.WhiteSmoke;
                        CountDown.ForeColor = Color.Black;
                    }
                    if (this.CountDown.Text != countdown.ToString(@"hh\:mm\:ss"))
                    {
                        this.setText(this.CountDown, countdown.ToString(@"hh\:mm\:ss"));
                    }
                    if (this.CountUp.Text != countup.ToString(@"hh\:mm\:ss"))
                    {
                        this.setText(this.CountUp, countup.ToString(@"hh\:mm\:ss"));
                    }
                    if (commitNow)
                    {
                        plcAuto.writeToPLC();
                        commitNow = false;
                    }
                    else
                    {
                        Thread.Sleep(500);
                    }
                }
            });
            await t;
            //after stop connections
            if (timerPLC.Enabled) timerPLC.Stop();
            for (int i = 0; i < plcAuto.PLCData.Count; i++)
            {
                plcAuto.PLCData[i].Value = 0;
            }
            plcAuto.PlcAutoLog.Enqueue(messages["DISCONN"]);
            setText(Quality, "");
            setEnabled(ConnectButton, true);
            setEnabled(DisconnectButton, false);
        }
        private void refreshPLC(object sender, EventArgs args)
        {
            if (Interlocked.Exchange(ref inTimer, 1) == 0)
            {
                if (_isStarting)
                {
                    if (IntData.IsOPCConnected)
                    {
                        plcAuto.writeToPLC();
                        plcAuto.PlcAutoLog.Enqueue(plcAuto.WaitEST ? messages["WAITEST"] : messages["STB"]);
                        Thread.Sleep(1000);
                    }
                }
                Interlocked.Exchange(ref inTimer, 0);
            }
        }
        private static string StrLenLimit(string str, int limitation)
        {
            int subLen = (limitation / 2) - 2;
            return str.Length > limitation ? str.Substring(0, subLen) + "...." + str.Substring(str.Length - subLen, subLen) : str;
        }
        private void setStatusLabel(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.statusBar.InvokeRequired)
            {
                SetStatusLabelCallback d = new SetStatusLabelCallback(setStatusLabel);
                if (this != null && !this.IsDisposed) this.Invoke(d, new object[] { text });
            }
            else
            {
                if (this != null && !this.IsDisposed) this.statusLabel.Text = text;
                Thread.Sleep(100);
                if (this != null && !this.IsDisposed) this.statusBar.Refresh();
            }
        }
        private void setText(Control control, string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (control.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(setText);
                this.Invoke(d, new object[] { control, text });
            }
            else
            {
                control.Text = text;
                control.Refresh();
            }
        }
        private void setEnabled(Control control, bool state)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (control.InvokeRequired)
            {
                SetEnableCallback d = new SetEnableCallback(setEnabled);
                if (this != null && !this.IsDisposed) this.Invoke(d, new object[] { control, state });
            }
            else
            {
                if (this != null && !this.IsDisposed) control.Enabled = state;
                if (this != null && !this.IsDisposed) control.Refresh();
            }
        }
        private void Panel_Load(object sender, EventArgs e)
        {
            plcAuto = new PLCAutomation(OPCName.Text, PLCName.Text);
        }
        private async void Start_Click(object sender, EventArgs e)
        {
            this.ControlBox = false;
            progressBar1.Visible = true;
            _isStarting = true;
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            GlobalLogStart();
            plcAuto.PlcAutoLog.Enqueue(messages["OPENCONN"]);
            plcAuto.PLCData = await Task.Run(() => initializing());
            progressBar1.Visible = false;
            var source = new BindingSource(plcAuto.PLCData, null);
            PLCDataGrid.DataSource = source;
            var t1 = Task.Run(() =>
            {
                while (_isStarting)
                {
                    if (PLCSimulatorSwitch.fetch)
                    {
                        if (PLCSimulatorSwitch.isSwitch)
                        {
                            ConnectButton_Click(sender, e);
                            PLCSimulatorSwitch.fetch = false;
                        }
                        else
                        {
                            DisconnectButton_Click(sender, e);
                            PLCSimulatorSwitch.fetch = false;
                        }
                    }

                    if (OPC2MQSwitch.fetch)
                    {
                        if (OPC2MQSwitch.isSwitch) OPCClientStart_Click(sender, e);
                        OPC2MQSwitch.fetch = false;
                        _opcState = OPC2MQSwitch.isSwitch;
                    }

                    if (MQ2DBSwitch.fetch)
                    {
                        if (!_q2DBState) Queue2DBStart_Click(sender, e);
                        MQ2DBSwitch.fetch = false;
                        _q2DBState = MQ2DBSwitch.isSwitch;
                    }

                    if (DSPbatchSwitch.fetch)
                    {
                        if (!_dspBatchState) DSPBatchStart_Click(sender, e);
                        DSPbatchSwitch.fetch = false;
                        _dspBatchState = DSPbatchSwitch.isSwitch;
                    }
                    PLCSimulatorSwitch.value = _isConnecting ? 1 : 0;
                    OPC2MQSwitch.value = _opcState ? 1 : 0;
                    MQ2DBSwitch.value = _q2DBState ? 1 : 0;
                    DSPbatchSwitch.value = _dspBatchState ? 1 : 0;
                    Thread.Sleep(100);
                }
            });
            await t1;
            this.ControlBox = true;
            progressBar1.Visible = false;
        }
        private void Stop_Click(object sender, EventArgs e)
        {
            DisconnectButton_Click(sender, e);
            DSPBatchStop_Click(sender, e);
            OPCClientStop_Click(sender, e);
            Queue2DBStop_Click(sender, e);
            toggleStartButtons(false);
            setEnabled(this.Stop, false);
            tabControl.SelectedIndex = 0;
        }
        private async void ConnectButton_Click(object sender, EventArgs e)
        {
            _isConnecting = true;
            refreshControls();
            timerPLC = new System.Timers.Timer(1000)
            {
                AutoReset = true,
                Enabled = true
            };
            timerPLC.Elapsed += new System.Timers.ElapsedEventHandler(refreshPLC);
            timerPLC.Start();
            setEnabled(ConnectButton, false);
            setEnabled(DisconnectButton, true);
            setEnabled(PLCDashboard, true);
            var t1 = Task.Run(() =>
            {
                while (_isConnecting)
                {
                    plcAuto.readFromPLC();
                    Thread.Sleep(1000);
                }
            });
            await t1;
            timerPLC.Stop();
            setEnabled(PLCDashboard, false);
        }
        private void Add5_Click(object sender, EventArgs e)
        {
            plcAuto.addCountUpMinutes(5);
        }
        private void Minus5_Click(object sender, EventArgs e)
        {
            plcAuto.addCountDownMinute(-5);
        }
        private void DisconnectButton_Click(object sender, EventArgs e)
        {
            _isConnecting = false;
        }
        private void ToggleDecode_CheckedChanged(object sender, EventArgs e)
        {
            plcAuto.isDecodeDecimal = ToggleDecode.Checked;
            if (plcAuto.isDecodeDecimal)
            {
                foreach (SwitchControl item in this.PLCDashboard.Controls.OfType<SwitchControl>().Where(item => indicators.Values.Contains(item.Name)))
                {
                    item.DisableSwitch();
                }
            }
            else
            {
                foreach (SwitchControl item in this.PLCDashboard.Controls.OfType<SwitchControl>().Where(item => indicators.Values.Contains(item.Name)))
                {
                    item.EnableSwitch();
                }
            }
        }
        private void Panel_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.DisconnectButton_Click(sender, e);
        }
        private async void DSPBatchStart_Click(object sender, EventArgs e)
        {
            _dspBatchState = true;
            _dspBatchSubState = true;
            setEnabled(DSPBatchStart, false);
            setEnabled(DSPBatchStop, true);
            setEnabled(DSPBatchControl, false);
            strSendPort = SendPort.Text;
            strServerIP = ServerIP.Text;
            strMailFrom = MailFrom.Text;
            strMailTo = MailTo.Text;
            strMailLogin = MailSMTPUserId.Text;
            strMailPort = Convert.ToInt32(MailSMTPort.Text);
            strMailHost = MailSMTPServer.Text;
            strTimerSecond = TimerSecond.Text;
            strConnectionSource = "Provider=SQLNCLI11;Persist Security Info=False;Data Source=" + this.DBAddr.Text + ";User ID=" + this.DBUser.Text + ";pwd=" + this.DBPwd.Text + ";Initial Catalog=GNA_IESO";
            strConnectionTarget = "Provider=SQLNCLI11;Persist Security Info=False;Data Source=" + this.DBAddr.Text + ";User ID=" + this.DBUser.Text + ";pwd=" + this.DBPwd.Text + ";Initial Catalog=MSS";
            strWinMsgSendInd = WinMsgSendInd.Text;
            strActionLogPath = ActionLogPath.Text;
            strActionLogDeleteDays = ActionLogDeleteDays.Text;
            strPlcHeartBeatMinute = PlcHeartBeatMinute.Text;
            isSendMail = isSendEmail.Checked;
            isSSLMail = isEmailSSL.Checked;
            strMailPass = emailPassword.Text;
            batch.OnStart();
            var t = Task.Run(() =>
            {
                string text = "";
                while (_dspBatchSubState || dspBatchLog.Count > 0)
                {
                    if (!_dspBatchState && _dspBatchSubState)
                    {
                        batch.OnStop();
                        _dspBatchSubState = false;
                    }
                    if (dspBatchLog.Count > 0) text = dspBatchLog.Dequeue() + Environment.NewLine + text;
                    if (text.Split('\r').Length > 22)
                    {
                        text = text.Remove(text.LastIndexOf(Environment.NewLine));
                    }
                    if (text != DSPBatchLog.Text)
                    {
                        setText(DSPBatchLog, text);
                    }
                    else
                    {
                        Thread.Sleep(300);
                    }

                }
            });
            var t1 = Task.Run(() =>
            {
                string text = "";
                while (_dspBatchSubState || dspBatchMessage.Count > 0)
                {
                    if (dspBatchMessage.Count > 0) text = dspBatchMessage.Dequeue() + Environment.NewLine + text;
                    if (text.Split('\r').Length > 20)
                    {
                        text = text.Remove(text.LastIndexOf(Environment.NewLine));
                    }
                    if (text != SocketMessageList.Text)
                    {
                        setText(SocketMessageList, text);
                    }
                    else
                    {
                        Thread.Sleep(300);
                    }

                }
            });
            var t2 = Task.Run(() =>
            {
                string text = "";
                while (_dspBatchSubState || dspBatchMail.Count > 0)
                {
                    if (dspBatchMail.Count > 0) text = dspBatchMail.Dequeue() + Environment.NewLine + text;
                    if (text.Split('\r').Length > 20)
                    {
                        text = text.Remove(text.LastIndexOf(Environment.NewLine));
                    }
                    if (text != EmailList.Text)
                    {
                        setText(EmailList, text);
                    }
                    else
                    {
                        Thread.Sleep(300);
                    }

                }
            });
            await t;
            await t1;
            await t2;
            setEnabled(DSPBatchStart, true);
            setEnabled(DSPBatchStop, false);
            setEnabled(DSPBatchControl, true);
        }
        private void DSPBatchStop_Click(object sender, EventArgs e)
        {
            _dspBatchState = false;
        }
        private async void OPCClientStart_Click(object sender, EventArgs e)
        {
            _opcState = true;
            _opcSubState = true;
            string _connOPC = "Data Source='" + this.DBAddr.Text + "'; Initial Catalog='OPC2DBMS';User id='" + this.DBUser.Text + "'; Password='" + this.DBPwd.Text + "';";
            string _connMSS = "Data Source='" + this.DBAddr.Text + "'; Initial Catalog='MSS';User id='" + this.DBUser.Text + "'; Password='" + this.DBPwd.Text + "';";

            OPC2Queue _ws = new OPC2Queue(_connOPC, _connMSS);
            Queue<string> opc2queue = new Queue<string>();
            setEnabled(OPCClientStop, true);
            setEnabled(OPCClientStart, false);
            var t = Task.Run(() =>
            {
                try
                {
                    opc2queue.Enqueue("Application starting...");
                    if (opc2queue.Count > 20) opc2queue.Dequeue();
                    _ws.Onstart();
                    while (_opcState)
                    {
                        if (_ws != null)
                        {
                            for (Logging info = _ws.log.FetchInformation(); info != null; info = _ws.log.FetchInformation())
                            {
                                opc2queue.Enqueue(info.Function + ": CMD" + StrLenLimit(info.Command, 100) + "\t MSG:" + info.Information);
                                if (opc2queue.Count > 20) opc2queue.Dequeue();
                            }
                            for (Logging warn = _ws.log.FetchWarning(); warn != null; warn = _ws.log.FetchWarning())
                            {
                                opc2queue.Enqueue(warn.Function + ": CMD" + StrLenLimit(warn.Command, 100) + "\t MSG:" + warn.Warning);
                                if (opc2queue.Count > 20) opc2queue.Dequeue();
                            }
                            for (Logging fatl = _ws.log.FetchWarning(); fatl != null; fatl = _ws.log.FetchFatal())
                            {
                                opc2queue.Enqueue(fatl.Function + ": CMD" + StrLenLimit(fatl.Command, 100) + "\t MSG:" + fatl.Exception.Message);
                                if (opc2queue.Count > 20) opc2queue.Dequeue();
                            }
                        }
                        Thread.Sleep(100);
                    }
                    opc2queue.Enqueue("The application is stoped.");
                    _ws.OnStop();
                }
                catch (Exception ex)
                {
                    opc2queue.Enqueue(ex.Message + "The application failed to start.");
                    if (opc2queue.Count > 20) opc2queue.Dequeue();
                }
            });
            var t1 = Task.Run(() =>
            {
                string text = "";
                while (_opcSubState || opc2queue.Count > 0)
                {
                    if (opc2queue.Count > 0) text = opc2queue.Dequeue() + Environment.NewLine + text;
                    if (text.Split('\r').Length > 20)
                    {
                        text = text.Remove(text.LastIndexOf(Environment.NewLine));
                    }
                    if (text != OPCClientLog.Text) setText(OPCClientLog, text);
                    Thread.Sleep(10);
                }
            });
            await t;
            _opcSubState = false;
            setEnabled(OPCClientStop, false);
            setEnabled(OPCClientStart, true);
        }
        private void OPCClientStop_Click(object sender, EventArgs e)
        {
            _opcState = false;
        }
        private async void Queue2DBStart_Click(object sender, EventArgs e)
        {

            _q2DBState = true;
            _q2DBSubState = true;
            string _connMSS = "Data Source='" + this.DBAddr.Text + "'; Initial Catalog='MSS';User id='" + this.DBUser.Text + "'; Password='" + this.DBPwd.Text + "';";

            Queue2DB _q2DB = new Queue2DB(_connMSS);
            Queue<string> queue2DBLog = new Queue<string>();
            setEnabled(Queue2DBStop, true);
            setEnabled(Queue2DBStart, false);
            var t = Task.Run(() =>
            {
                try
                {
                    queue2DBLog.Enqueue("Application starting...");
                    if (queue2DBLog.Count > 20) queue2DBLog.Dequeue();
                    _q2DB.OnStart();
                    while (_q2DBState || _q2DB.ResultQueue.Count > 0)
                    {
                        if (_q2DB.ResultQueue.Count > 0)
                        {
                            queue2DBLog.Enqueue(_q2DB.ResultQueue.Dequeue().Command);
                        }
                        else
                        {
                            Thread.Sleep(300);
                        }

                    }
                    queue2DBLog.Enqueue("The application stopped.");
                    _q2DB.OnStop();
                }
                catch (Exception ex)
                {
                    queue2DBLog.Enqueue(ex.Message + "The application failed to start.");
                    if (queue2DBLog.Count > 20) queue2DBLog.Dequeue();
                }
            });
            var t1 = Task.Run(() =>
            {
                string text = "";
                while (_q2DBSubState || queue2DBLog.Count > 0)
                {
                    if (queue2DBLog.Count > 0) text = queue2DBLog.Dequeue() + Environment.NewLine + text;
                    if (text.Split('\r').Length > 20)
                    {
                        text = text.Remove(text.LastIndexOf(Environment.NewLine));
                    }
                    if (text != OPCClientLog.Text) setText(Queue2DBLog, text);
                    Thread.Sleep(300);
                }
            });
            await t;
            _q2DBSubState = false;
            setEnabled(Queue2DBStop, false);
            setEnabled(Queue2DBStart, true);
        }
        private void Queue2DBStop_Click(object sender, EventArgs e)
        {
            _q2DBState = false;
        }
    }
}
