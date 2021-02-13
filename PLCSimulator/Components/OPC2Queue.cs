using System;
using System.Timers;
using PLCTools.Common;
using PLCTools.Service;
using PLCTools.Models;

namespace PLCTools.Components
{
    public class OPC2Queue
    {
        public Logging log = new Logging();
        private Timer timer;
        private readonly TagHandler Tags = new TagHandler();
        private readonly EventsHandler Events = new EventsHandler();
        public OPC2Queue(string _connOPC = "", string _connMSS = "")
        {
            if (_connOPC != "") IntData.CfgConn = _connOPC;
            if (_connMSS != "") IntData.DestConn = _connMSS;
        }
        private void Initialization()
        {
            IntData.InitializeData();
            Tags.Init();
            Events.Init();
            MQHandler queue = new MQHandler();
            queue.CreateQueue();
            timer = new Timer(IntData.timeInterval)
            {
                AutoReset = true,
                Enabled = true
            };
            timer.Elapsed += new ElapsedEventHandler(SetTaskAtFixedTime);
            timer.Start();
        }
        private void SetTaskAtFixedTime(object sender, EventArgs e)
        {
            Events.Refresh(Tags);
            if (!IntData.IsOPCConnected)
            {
                Initialization();
            }
        }
        public void Onstart()
        {
            try
            {
                Initialization();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + " Initialization failed");
            }
        }
        public void OnStop()
        {
            if (timer != null) timer.Stop();
        }
    }
}