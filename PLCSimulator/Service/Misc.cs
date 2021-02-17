using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PLCTools.Service
{
    public class Misc
    {
        private delegate void SetTextCallback(Control control, string text);
        internal bool isConnecting { get; set; } = false;
        internal Queue<string> Log { get; set; } = new Queue<string>();
        internal string StrLenLimit(string str, int limitation)
        {
            int subLen = (limitation / 2) - 2;
            return str.Length > limitation ? str.Substring(0, subLen) + "...." + str.Substring(str.Length - subLen, subLen) : str;
        }
        internal void LogToTextBox(Control textbox, int strLength)
        {

            Task.Run(() =>
            {
                while (Log.Count > 0 || isConnecting)
                {
                    if (Log.Count > 0)
                    {
                        string text = "";
                        text = Log.Dequeue() + Environment.NewLine + textbox.Text;
                        if (text.Split('\r').Length > strLength)
                        {
                            text = text.Remove(text.LastIndexOf(Environment.NewLine));
                        }
                        if (text != textbox.Text) setText(textbox, text);
                    }
                    Thread.Sleep(10);
                }
            });
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
        internal void OnStop()
        {
            Log.Enqueue("Application Stopping...");
            isConnecting = false;
        }
    }
}
