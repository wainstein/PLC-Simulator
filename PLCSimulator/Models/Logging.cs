using System;
using System.Diagnostics;
using PLCTools.Common;

namespace PLCTools.Service
{
    public class Logging
    {
        public string Function { get; set; } = "";
        public string Command { get; set; } = "";
        public string Information { get; set; } = "";
        public string Warning { get; set; } = "";
        public Exception Exception { get; set; } = null;
        public Logging(string cmd = "", string func = "")
        {
            if (func == "") func = new StackFrame(1, true).GetMethod().Name;
            this.Function = func;
            if (cmd != "") this.Command = cmd;
        }
        public Boolean Success(string infor = "Successful")
        {
            this.Information = infor;
            if (IntData.InforQ.Count > 200)
            {
                IntData.InforQ.Dequeue();
            }
            IntData.InforQ.Enqueue(this);
            return true;
        }
        public Boolean Fatal(Exception ex)
        {
            this.Exception = ex;
            if (IntData.ErrorQ.Count > 200)
            {
                IntData.ErrorQ.Dequeue();
            }
            IntData.ErrorQ.Enqueue(this);
            return true;
        }
        public Boolean Warn(string infor)
        {
            this.Information = infor;
            if (IntData.WarningQ.Count > 200)
            {
                IntData.WarningQ.Dequeue();
            }
            IntData.WarningQ.Enqueue(this);
            return true;
        }
        public Logging FetchInformation()
        {
            if (IntData.InforQ.Count > 0)
            {
                return IntData.InforQ.Dequeue();
            }
            return null;
        }
        public Logging FetchWarning()
        {
            if (IntData.WarningQ.Count > 0)
            {
                return IntData.WarningQ.Dequeue();
            }
            return null;
        }
        public Logging FetchFatal()
        {
            if (IntData.ErrorQ.Count > 0)
            {
                return IntData.ErrorQ.Dequeue();
            }
            return null;
        }
    }
}