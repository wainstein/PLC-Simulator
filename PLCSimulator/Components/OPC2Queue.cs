using System;
using System.Timers;
using PLCTools.Common;
using PLCTools.Service;
using System.Threading.Tasks;
using System.Threading;
using System.Collections.Generic;

namespace PLCTools.Components
{
    public class OPC2Queue : Misc
    {
        private Logging log { get; set; } = new Logging();
        private int eventTimer;
        private readonly TagHandler Tags = new TagHandler();
        private readonly EventsHandler Events = new EventsHandler();
        public OPC2Queue(string _connOPC = "", string _connMSS = "")
        {
            if (_connOPC != "") IntData.CfgConn = _connOPC;
            if (_connMSS != "") IntData.DestConn = _connMSS;
        }
        public async void Onstart()
        {
            isConnecting = true;
            MQHandler queue = new MQHandler();
            queue.CreateQueue();
            Tags.Init();
            Events.Init();
            var t = Task.Run(() =>
            {
                while (isConnecting)
                {
                    Task.Run(() =>
                    {
                        if (Interlocked.Exchange(ref eventTimer, 1) == 0)
                        {
                            Events.Refresh(Tags);
                            Interlocked.Exchange(ref eventTimer, 0);
                        }
                    });
                    Thread.Sleep(1000);
                }
            });
            var t1 = Task.Run(() =>
            {
                try
                {
                    Log.Enqueue("Application starting...");
                    if (Log.Count > 20) Log.Dequeue();
                    while (isConnecting)
                    {
                        for (Logging info = log.FetchInformation(); info != null; info = log.FetchInformation())
                        {
                            Log.Enqueue(info.Function + ": CMD" + StrLenLimit(info.Command, 100) + "\t MSG:" + info.Information);
                            if (Log.Count > 20) Log.Dequeue();
                        }
                        for (Logging warn = log.FetchWarning(); warn != null; warn = log.FetchWarning())
                        {
                            Log.Enqueue(warn.Function + ": CMD" + StrLenLimit(warn.Command, 100) + "\t MSG:" + warn.Warning);
                            if (Log.Count > 20) Log.Dequeue();
                        }
                        for (Logging fatl = log.FetchWarning(); fatl != null; fatl = log.FetchFatal())
                        {
                            Log.Enqueue(fatl.Function + ": CMD" + StrLenLimit(fatl.Command, 100) + "\t MSG:" + fatl.Exception.Message);
                            if (Log.Count > 20) Log.Dequeue();
                        }
                    }
                    Thread.Sleep(100);
                    Log.Enqueue("The application is stoped.");
                }
                catch (Exception ex)
                {
                    Log.Enqueue(ex.Message + "The application failed to start.");
                    if (Log.Count > 20) Log.Dequeue();
                }
            });
            await t;
            Events.Dispose();
            Tags.Dispose();
        }
    }
}