using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using PLCTools.Common;
using PLCTools.Models;
using Timer = System.Timers.Timer;

namespace PLCTools.Components
{
    class PLCAutomation
    {
        internal BindingList<OPCItems> PLCData = new BindingList<OPCItems>();
        internal string ServerName { get; set; }
        internal string PLCName { get; set; }
        internal bool isDecodeDecimal { get; set; } = true;
        internal Queue<string> PlcAutoLog { get; set; } = new Queue<string>();
        internal string OverallQuality { get; set; }
        internal bool WaitEST { get; set; } = false;
        internal bool isCountDown { get; set; } = false;
        internal bool isCountUp { get; set; } = false;
        internal DateTime countingDownTime { get; set; } = DateTime.Now;
        internal DateTime countingUpTime { get; set; } = DateTime.Now;
        internal Queue<string> transmitQueue = new Queue<string>();
        private static int inTimer = 0;
        private Dictionary<string, OPCItems> writeTaskDic { get; set; } = new Dictionary<string, OPCItems>();
        public PLCAutomation(string servername, string plcname)
        {
            this.ServerName = servername;
            this.PLCName = plcname;
        }
        internal void queuePLCWrites(int value, string tag)
        {
            OPCItems Item = getTagItem(tag);
            if (Convert.ToInt32(Item.Value) != value)
            {
                Item.Value = value;
                if (writeTaskDic.ContainsKey(Item.Tag))
                {
                    writeTaskDic[Item.Tag] = Item;
                }
                else
                {
                    writeTaskDic.Add(Item.Tag, Item);
                }
            }
        }
        internal int getTagValue(string tag)
        {
            for (int i = 0; i < PLCData.Count; i++)
            {
                if (PLCData[i].Tag == tag) return Convert.ToInt32(PLCData[i].Value);
            }
            return -1;
        }
        internal OPCItems getTagItem(string tag)
        {
            for (int i = 0; i < PLCData.Count; i++)
            {
                if (PLCData[i].Tag == tag) return PLCData[i];
            }
            return null;
        }
        internal void readFromPLC()
        {
            using (OPCController oPCGroups = new OPCController(ServerName, "PLCAutomationRead", PLCName))
            {
                PlcAutoLog.Enqueue(WaitEST ? Panel.messages["WAITEST"] : Panel.messages["CONNOPC"]);
                if (oPCGroups.ServerName.Length > 1)
                {
                    oPCGroups.GetData(ref PLCData);
                    string Quality = "Good";
                    for (int i = 0; i < PLCData.Count; i++)
                    {
                        OPCItems oPCItems = oPCGroups.GetTagItem(PLCData[i].Tag);
                        if (oPCItems.Quality < 192) Quality = "Bad";
                        PLCData[i].Value = oPCItems.Value;
                        PLCData[i].Quality = oPCItems.Quality;
                    }
                    OverallQuality = Quality;
                    PlcAutoLog.Enqueue(WaitEST ? Panel.messages["WAITEST"] : Panel.messages["OPCREAD"]);
                }
            }
        }
        internal void writeToPLC()
        {
            preparedWritesQueue();
            List<OPCItems> taskList = new List<OPCItems>();
            foreach (OPCItems item in writeTaskDic.Values) taskList.Add(item);
            writeTaskDic = new Dictionary<string, OPCItems>();
            if (taskList.Count > 0)
            {
                using (OPCController oPCGroup_write = new OPCController(ServerName, "Write", PLCName))
                {
                    oPCGroup_write.PutData(taskList);
                }
            }
        }
        private void refreshPLC(object sender, EventArgs args)
        {
            if (Interlocked.Exchange(ref inTimer, 1) == 0)
            {
                if (IntData.IsOPCConnected)
                {
                    writeToPLC();
                    PlcAutoLog.Enqueue(WaitEST ? Panel.messages["WAITEST"] : Panel.messages["STB"]);
                    Thread.Sleep(1000);
                }
                Interlocked.Exchange(ref inTimer, 0);
            }
        }
        private void preparedWritesQueue()
        {
            if (getTagValue("DSP_READ_HEARTBEAT") != getTagValue("DSP_WRITE_HEARTBEAT") + 1)
            {
                queuePLCWrites(getTagValue("DSP_WRITE_HEARTBEAT") + 1, "DSP_READ_HEARTBEAT");
            }

            if (Convert.ToInt32(getTagValue("DSP_READ_POWER_ONOFF")) == 0)
            {

                TimeSpan spanUp = DateTime.Now - countingUpTime;
                TimeSpan spanDown = countingDownTime - DateTime.Now;
                if (isCountUp)
                {
                    isCountDown = false;
                    WaitEST = false;
                    var min = Convert.ToInt16(spanUp.TotalMinutes);
                    var sec = spanUp.Seconds;
                    queuePLCWrites(min, "DSP_CUP_MIN");
                    queuePLCWrites(sec, "DSP_CUP_SEC");
                    if (spanUp.TotalSeconds >= Convert.ToInt32(getTagValue("DSP_DISPATCH_ON_PWR_OFF")) * 60)
                    {
                        isCountUp = false;
                        WaitEST = true;
                    }
                }
                if (WaitEST)
                {
                    isCountUp = false;
                    isCountDown = false;
                    int estStartMin = Convert.ToInt32(getTagValue("DSP_EST_START_MIN"));
                    if (estStartMin == 0)
                    {
                        isCountUp = false;
                        countingDownTime = DateTime.Now;
                        queuePLCWrites(0, "DSP_CDOWN_MIN");
                        queuePLCWrites(0, "DSP_CDOWN_SEC");
                        isCountDown = false;
                        PlcAutoLog.Enqueue(Panel.messages["WAITEST"]);
                    }
                    else
                    {
                        countingDownTime = DateTime.Now + new TimeSpan(TimeSpan.TicksPerMinute * estStartMin);
                        var min = Convert.ToInt16(spanDown.TotalMinutes);
                        var sec = spanDown.Seconds;
                        queuePLCWrites(min, "DSP_CDOWN_MIN");
                        queuePLCWrites(sec, "DSP_CDOWN_SEC");
                        WaitEST = false;
                        isCountDown = true;
                        isCountUp = false;
                    }

                }
                else
                {
                    if (isCountDown)
                    {
                        isCountUp = false;
                        WaitEST = false;
                        var min = Convert.ToInt16(spanDown.TotalMinutes);
                        var sec = spanDown.Seconds;
                        queuePLCWrites(min, "DSP_CDOWN_MIN");
                        queuePLCWrites(sec, "DSP_CDOWN_SEC");
                        if (spanDown.TotalSeconds <= 0)
                        {
                            min = 0;
                            sec = 0;
                            queuePLCWrites(min, "DSP_CDOWN_MIN");
                            queuePLCWrites(sec, "DSP_CDOWN_SEC");
                            isCountUp = false;
                            WaitEST = true;
                            isCountDown = false;
                        }
                        if (Convert.ToInt32(getTagValue("DSP_DISPATCH_ON")) == 1 && Convert.ToInt32(getTagValue("DSP_DISPATCH_OFF")) == 0)
                        {
                            queuePLCWrites(0, "DSP_CUP_MIN");
                            queuePLCWrites(0, "DSP_CUP_SEC");
                            queuePLCWrites(0, "DSP_CDOWN_MIN");
                            queuePLCWrites(0, "DSP_CDOWN_SEC");
                            countingUpTime = DateTime.Now;
                            WaitEST = false;
                            isCountDown = false;
                            isCountUp = true;
                        }
                    }
                }
                if (!isCountUp && !isCountDown && !WaitEST)
                {
                    if (!(spanDown.TotalSeconds <= 0 && spanUp.TotalSeconds >= Convert.ToInt32(getTagValue("DSP_DISPATCH_ON_PWR_OFF")) * 60))
                    {
                        isCountUp = true;
                        isCountDown = false;
                        WaitEST = false;
                        countingUpTime = DateTime.Now;
                    }
                }
            }
            else
            {
                countingDownTime = DateTime.Now;
                countingUpTime = DateTime.Now;
                queuePLCWrites(0, "DSP_CUP_MIN");
                queuePLCWrites(0, "DSP_CUP_SEC");
                queuePLCWrites(0, "DSP_CDOWN_MIN");
                queuePLCWrites(0, "DSP_CDOWN_SEC");
                isCountUp = false;
                isCountDown = false;
            }

            if (isDecodeDecimal)
            {
                decodeDecimalWrite(getTagValue("DSP_DECIMAL_WRITE"));
            }
        }
        private void decodeDecimalWrite(int decimalWrite)
        {
            if (decimalWrite < 0)
            {
                queuePLCWrites(1, "DSP_AUTO_DIAL1");
                decimalWrite += 32768;
            }
            else
            {
                queuePLCWrites(0, "DSP_AUTO_DIAL1");
            }
            string decodeString = Convert.ToString(decimalWrite, 2);

            int misBit = 15 - decodeString.Length;

            for (int i = 0; i < misBit; i++)
            {
                decodeString = "0" + decodeString;
            }
            for (int i = 0; i < decodeString.Length; i++)
            {
                if (Panel.indicators.ContainsKey(i)) queuePLCWrites(Convert.ToInt32(decodeString[i].ToString()), Panel.indicators[i]);
            }
        }
        internal void addCountUpMinutes(int minutes)
        {
            queuePLCWrites(Convert.ToInt32(getTagValue("DSP_CUP_MIN")) + minutes, "DSP_CUP_MIN");
            writeToPLC();
            countingUpTime = countingUpTime.AddMinutes(-minutes);
        }
        internal void addCountDownMinute(int minutes)
        {
            queuePLCWrites(Convert.ToInt32(getTagValue("DSP_CDOWN_MIN")) + minutes, "DSP_CDOWN_MIN");
            writeToPLC();
            countingDownTime = countingDownTime.AddMinutes(minutes);
        }
    }
}
