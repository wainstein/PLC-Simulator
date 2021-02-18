using PLCTools.Common;
using PLCTools.Components;
using PLCTools.Service;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PLCTools
{
    public partial class Panel : Form
    {
        delegate void SetStatusLabelCallback(string text);
        delegate void SetTextCallback(Control control, string text);
        delegate void SetEnableCallback(Control control, bool state);
        internal static Queue<string> PlcAutoLog { get; set; } = new Queue<string>();
        private bool _isStarting { get; set; } = false;
        private DspBatch batch = new DspBatch();
        private OPC2Queue _ws;
        private Queue2DB _q2DB;
        private PLCAutomation _plcAuto;
        private string _connOPC { get; set; }
        private string _connMSS { get; set; }
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
        private static Dictionary<string, Control> leds { get; set; } = new Dictionary<string, Control>();
        private static bool[] busyLeds = new bool[17];
        public Panel()
        {
            InitializeComponent();
            foreach (Control control in Indicators.Controls)
            {
                leds.Add(control.Name, control);
            }
            foreach (SwitchControl item in this.PLCDashboard.Controls.OfType<SwitchControl>().Where(item => indicators.Values.Contains(item.Name)))
            {
                item.DisableSwitch();
            }
        }
        private BindingList<OPCItems> initPLCData()
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
                        PlcAutoLog.Enqueue(messages["RETRIVE"]);
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
                    }
                    PlcAutoLog.Enqueue(messages["DBSUCCESS"]);
                }
                catch (Exception ex)
                {
                    PlcAutoLog.Enqueue(messages["DBFAIL"]);
                    toggleStartButtons(false);
                    Console.WriteLine(ex);
                }
            }
            return plcData;
        }
        private void toggleStartButtons(bool state)
        {
            while (PlcAutoLog.Count > 0)
            {
                setStatusLabel(PlcAutoLog.Dequeue());
                Thread.Sleep(100);
            }
            _isStarting = state;
            setEnabled(this.Stop, state);
            setEnabled(this.SwitchPanel, state);
            setEnabled(this.tabControl, state);
            setEnabled(this.panelStart, !state);
            setEnabled(this.Start, !state);
        }
        private async void GlobalLogStart()
        {
            var t = Task.Run(() =>
             {
                 while (_isStarting || PlcAutoLog.Count > 0)
                 {
                     if (Quality.Text != _plcAuto.OverallQuality) setText(Quality, _plcAuto.OverallQuality);
                     if (PlcAutoLog.Count > 0)
                     {
                         setStatusLabel(PlcAutoLog.Dequeue());
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
        private void refreshIndicators(int idx)
        {
            var t = Task.Run(() =>
            {
                Control label = leds["L" + idx];
                Control transLed = leds["T" + idx];
                Control xcieveLed = leds["X" + idx];
                try
                {
                    busyLeds[idx] = true;
                    if (IntData.OPCControllers.Count >= idx)
                    {
                        OPCController item = IntData.OPCControllers[idx - 1];
                        if (item != null)
                        {
                            if (!label.Enabled)
                            {
                                setEnabled(label, true);
                                setEnabled(transLed, true);
                                setEnabled(xcieveLed, true);
                            }
                            if (item.GroupName != label.Text)
                            {
                                setText(label, item.GroupName);
                            }
                            if (transLed.BackColor == Color.DimGray)
                            {
                                if (item.TransactionFlag)
                                {
                                    transLed.BackColor = Color.Lime;
                                    Thread.Sleep(200);
                                }
                            }
                            else
                            {
                                if (!item.Transacting)
                                {
                                    item.TransactionFlag = false;
                                    transLed.BackColor = Color.DimGray;
                                    Thread.Sleep(50);
                                    for (int i = 0; i < item.TransationSum; i++)
                                    {
                                        if (item.Quality == "Good")
                                        {
                                            xcieveLed.BackColor = Color.Lime;
                                        }
                                        else
                                        {
                                            xcieveLed.BackColor = Color.Crimson;
                                        }
                                        Thread.Sleep(200 / item.TransationSum);
                                        xcieveLed.BackColor = Color.DimGray;
                                        Thread.Sleep(200 / item.TransationSum);
                                    }
                                }
                            }

                        }
                        else
                        {
                            if (xcieveLed.Text != "(EMPTY)")
                            {
                                setText(label, "(EMPTY)");
                                xcieveLed.BackColor = Color.DimGray;
                                transLed.BackColor = Color.DimGray;
                                setEnabled(label, false);
                                setEnabled(transLed, false);
                                setEnabled(xcieveLed, false);
                            }
                        }
                    }
                    else
                    {
                        if (label.Text != "(READY)")
                        {
                            if (transLed.BackColor == Color.Lime)
                            {
                                transLed.BackColor = Color.DimGray;
                                Thread.Sleep(50);
                                xcieveLed.BackColor = Color.Lime;
                                Thread.Sleep(200);
                                xcieveLed.BackColor = Color.DimGray;
                            }
                            if (label.Text != "Channel " + idx) Thread.Sleep(500);
                            setText(label, "(READY)");
                            setEnabled(label, false);
                            setEnabled(transLed, false);
                            setEnabled(xcieveLed, false);
                        }
                    }

                }
                catch
                {
                    if (label.Text != "ERROR")
                    {
                        setText(xcieveLed, "ERROR");
                        xcieveLed.BackColor = Color.DimGray;
                        transLed.BackColor = Color.DimGray;
                        setEnabled(label, false);
                        setEnabled(transLed, false);
                        setEnabled(xcieveLed, false);
                    }
                }
                finally
                {
                    busyLeds[idx] = false;
                }
            });
        }
        private async void refreshControls()
        {
            var t = Task.Run(() =>
            {
                bool commitNow = false;
                while (_plcAuto.isConnecting && _isStarting)
                {
                    foreach (SwitchControl item in this.PLCDashboard.Controls.OfType<SwitchControl>())
                    {
                        if (!item.fetch)
                        {
                            if (item.value != _plcAuto.getTagValue(item.Name) || item.Pending)
                            {
                                item.value = _plcAuto.getTagValue(item.Name);
                                item.quality = _plcAuto.getTagItem(item.Name) == null ? 0 : (_plcAuto.getTagItem(item.Name).Quality / 192.0);
                            }
                        }
                        else
                        {
                            _plcAuto.queuePLCWrites(item.value, item.Name);
                            Thread.Sleep(200);
                            item.fetch = false;
                            commitNow = true;
                        }
                        if (item.Text != _plcAuto.getTagValue(item.Name).ToString())
                        {
                            this.setText(item, _plcAuto.getTagValue(item.Name).ToString());
                        }
                    }
                    foreach (TextBox item in this.PLCDashboard.Controls.OfType<TextBox>())
                    {
                        if (item.Text != _plcAuto.getTagValue(item.Name).ToString())
                        {
                            this.setText(item, _plcAuto.getTagValue(item.Name).ToString());
                        }
                    }

                    TimeSpan countdown = new TimeSpan(0);
                    TimeSpan countup = new TimeSpan(0);
                    if (_plcAuto.isCountDown)
                    {
                        CountDown.BackColor = Color.RoyalBlue;
                        CountDown.ForeColor = Color.WhiteSmoke;
                        CountUp.BackColor = Color.WhiteSmoke;
                        CountUp.ForeColor = Color.Black;
                        countdown = _plcAuto.countingDownTime - DateTime.Now;
                    }
                    else if (_plcAuto.WaitEST)
                    {
                        CountDown.BackColor = Color.RoyalBlue;
                        CountDown.ForeColor = Color.WhiteSmoke;
                        CountUp.BackColor = Color.WhiteSmoke;
                        CountUp.ForeColor = Color.Black;
                    }
                    else if (_plcAuto.isCountUp)
                    {
                        CountUp.BackColor = Color.RoyalBlue;
                        CountUp.ForeColor = Color.WhiteSmoke;
                        CountDown.BackColor = Color.WhiteSmoke;
                        CountDown.ForeColor = Color.Black;
                        countup = DateTime.Now - _plcAuto.countingUpTime;
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
                        _plcAuto.Write();
                        commitNow = false;
                    }
                    else
                    {
                        Thread.Sleep(500);
                    }
                }
            });
            await t;
            foreach (SwitchControl item in this.PLCDashboard.Controls.OfType<SwitchControl>())
            {
                item.isSwitch = false;
                item.value = 0;
            }
            //after stop connections
            for (int i = 0; i < _plcAuto.PLCData.Count; i++)
            {
                _plcAuto.PLCData[i].Value = 0;
            }
            PlcAutoLog.Enqueue(messages["DISCONN"]);
            setText(Quality, "");
            setEnabled(ConnectButton, true);
            setEnabled(DisconnectButton, false);
        }
        private void setStatusLabel(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.statusBar.InvokeRequired)
            {
                SetStatusLabelCallback d = new SetStatusLabelCallback(setStatusLabel);
                if (this.statusBar != null && !this.statusBar.IsDisposed) this.statusBar.Invoke(d, new object[] { text });
            }
            else
            {
                if (this.statusBar != null && !this.statusBar.IsDisposed) this.statusLabel.Text = text;
                Thread.Sleep(100);
                if (this.statusBar != null && !this.statusBar.IsDisposed) this.statusBar.Refresh();
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
                control.Invoke(d, new object[] { control, text });
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
                if (control != null && !control.IsDisposed) control.Invoke(d, new object[] { control, state });
            }
            else
            {
                if (control != null && !control.IsDisposed) control.Enabled = state;
                if (control != null && !control.IsDisposed) control.Refresh();
            }
        }
        private void Panel_Load(object sender, EventArgs e)
        {
            _plcAuto = new PLCAutomation(OPCName.Text, PLCName.Text);
        }
        private async void Start_Click(object sender, EventArgs e)
        {
            this.ControlBox = false;
            progressBar1.Visible = true;
            _isStarting = true;
            GlobalLogStart();
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            PlcAutoLog.Enqueue(messages["OPENCONN"]);
            _plcAuto.PLCData = await Task.Run(() => initPLCData());
            var t = Task.Run(() =>
            {
                while (_isStarting)
                {
                    for (int i = 1; i <= 16; i++)
                    {
                        if (!busyLeds[i])
                        {
                            refreshIndicators(i);
                        }
                    }
                    Thread.Sleep(1000);
                }
            });
            IntData.InitializeData();
            var source = new BindingSource(_plcAuto.PLCData, null);
            PLCDataGrid.DataSource = source;
            _connOPC = "Data Source='" + this.DBAddr.Text + "'; Initial Catalog='OPC2DBMS';User id='" + this.DBUser.Text + "'; Password='" + this.DBPwd.Text + "';";
            _connMSS = "Data Source='" + this.DBAddr.Text + "'; Initial Catalog='MSS';User id='" + this.DBUser.Text + "'; Password='" + this.DBPwd.Text + "';";
            if (_ws == null) _ws = new OPC2Queue(_connOPC, _connMSS);
            if (_q2DB == null) _q2DB = new Queue2DB(_connMSS);
            toggleStartButtons(true);
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
                        _ws.isConnecting = OPC2MQSwitch.isSwitch;
                    }

                    if (MQ2DBSwitch.fetch)
                    {
                        if (!_q2DB.isConnecting) Queue2DBStart_Click(sender, e);
                        MQ2DBSwitch.fetch = false;
                        _q2DB.isConnecting = MQ2DBSwitch.isSwitch;
                    }
                    if (DSPbatchSwitch.fetch)
                    {
                        if (!batch.isConnecting) DSPBatchStart_Click(sender, e);
                        DSPbatchSwitch.fetch = false;
                        batch.isConnecting = DSPbatchSwitch.isSwitch;
                    }
                    if (PLCSimulatorSwitch.isSwitch != _plcAuto.isConnecting || PLCSimulatorSwitch.Pending)
                    {
                        PLCSimulatorSwitch.value = _plcAuto.isConnecting ? 1 : 0;
                        setEnabled(PLCDashboard, _plcAuto.isConnecting);
                        setEnabled(ConnectButton, !_plcAuto.isConnecting);
                        setEnabled(DisconnectButton, _plcAuto.isConnecting);
                    }
                    if (OPC2MQSwitch.isSwitch != _ws.isConnecting || OPC2MQSwitch.Pending)
                    {
                        OPC2MQSwitch.value = _ws.isConnecting ? 1 : 0;
                        setEnabled(OPCClientStop, _ws.isConnecting);
                        setEnabled(OPCClientStart, !_ws.isConnecting);
                    }
                    if (MQ2DBSwitch.isSwitch != _q2DB.isConnecting || MQ2DBSwitch.Pending)
                    {
                        MQ2DBSwitch.value = _q2DB.isConnecting ? 1 : 0;
                        setEnabled(Queue2DBStop, _q2DB.isConnecting);
                        setEnabled(Queue2DBStart, !_q2DB.isConnecting);
                    }
                    if (DSPbatchSwitch.isSwitch != batch.isConnecting || DSPbatchSwitch.Pending)
                    {
                        DSPbatchSwitch.value = batch.isConnecting ? 1 : 0;
                        setEnabled(DSPBatchStart, !batch.isConnecting);
                        setEnabled(DSPBatchStop, batch.isConnecting);
                        setEnabled(DSPBatchControl, !batch.isConnecting);
                    }
                    Thread.Sleep(200);
                }
            });
            progressBar1.Visible = false;
            await t1;
            this.ControlBox = true;            
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
        private void ConnectButton_Click(object sender, EventArgs e)
        {
            _plcAuto.OnStart();
            refreshControls();
        }
        private void Add5_Click(object sender, EventArgs e)
        {
            _plcAuto.addCountUpMinutes(5);
        }
        private void Minus5_Click(object sender, EventArgs e)
        {
            _plcAuto.addCountDownMinute(-5);
        }
        private void DisconnectButton_Click(object sender, EventArgs e)
        {
            _plcAuto.OnStop();
        }
        private void ToggleDecode_CheckedChanged(object sender, EventArgs e)
        {
            _plcAuto.isDecodeDecimal = ToggleDecode.Checked;
            if (_plcAuto.isDecodeDecimal)
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
            batch = new DspBatch()
            {
                isConnecting = true,
                _dspBatchSubState = true,
                strSendPort = SendPort.Text,
                strServerIP = ServerIP.Text,
                strMailFrom = MailFrom.Text,
                strMailTo = MailTo.Text,
                strMailLogin = MailSMTPUserId.Text,
                strMailPort = Convert.ToInt32(MailSMTPort.Text),
                strMailHost = MailSMTPServer.Text,
                strTimerSecond = TimerSecond.Text,
                strConnectionSource = "Provider=SQLNCLI11;Persist Security Info=False;Data Source=" + this.DBAddr.Text + ";User ID=" + this.DBUser.Text + ";pwd=" + this.DBPwd.Text + ";Initial Catalog=GNA_IESO",
                strConnectionTarget = "Provider=SQLNCLI11;Persist Security Info=False;Data Source=" + this.DBAddr.Text + ";User ID=" + this.DBUser.Text + ";pwd=" + this.DBPwd.Text + ";Initial Catalog=MSS",
                strWinMsgSendInd = WinMsgSendInd.Text,
                strActionLogPath = ActionLogPath.Text,
                strActionLogDeleteDays = ActionLogDeleteDays.Text,
                strPlcHeartBeatMinute = PlcHeartBeatMinute.Text,
                isSendMail = isSendEmail.Checked,
                isSSLMail = isEmailSSL.Checked,
                strMailPass = emailPassword.Text
            };
            batch.OnStartAsync();
            var t = Task.Run(() =>
            {
                string text = "";
                while (batch._dspBatchSubState || IntData.dspBatchLog.Count > 0)
                {
                    if (!batch.isConnecting && batch._dspBatchSubState)
                    {
                        batch.OnStop();
                        batch._dspBatchSubState = false;
                    }
                    if (IntData.dspBatchLog.Count > 0) text = IntData.dspBatchLog.Dequeue() + Environment.NewLine + text;
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
                while (batch._dspBatchSubState || batch.dspBatchMessage.Count > 0)
                {
                    if (batch.dspBatchMessage.Count > 0) text = batch.dspBatchMessage.Dequeue() + Environment.NewLine + text;
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
                while (batch._dspBatchSubState || IntData.dspBatchMail.Count > 0)
                {
                    if (IntData.dspBatchMail.Count > 0) text = IntData.dspBatchMail.Dequeue() + Environment.NewLine + text;
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
        }
        private void DSPBatchStop_Click(object sender, EventArgs e)
        {
            batch.OnStop();
        }
        private void OPCClientStart_Click(object sender, EventArgs e)
        {
            _ws.Onstart();
            _ws.Log.Enqueue("Application starting...");
            _ws.LogToTextBox(OPCClientLog, 20);
        }
        private void OPCClientStop_Click(object sender, EventArgs e)
        {
            _ws.OnStop();
        }
        private void Queue2DBStart_Click(object sender, EventArgs e)
        {
            _q2DB.OnStart();
            _q2DB.Log.Enqueue("Application starting...");
            _q2DB.LogToTextBox(Queue2DBLog, 20);
        }
        private void Queue2DBStop_Click(object sender, EventArgs e)
        {
            _q2DB.OnStop();
        }
    }
}