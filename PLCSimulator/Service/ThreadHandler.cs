using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PLCTools.Service
{
    class ThreadHandler
    {
        public static Boolean ThreadLocked { get; set; } = false;
        public static int ThreadLocker(int timeInterval = 50)
        {
            int time = 0;
            while (ThreadLocked)
            {
                Thread.Sleep(timeInterval);
                time++;
            }
            return time * timeInterval;
        }
        public static void LockThread() { ThreadLocked = true; }
        public static void ReleaseThread() { ThreadLocked = false; }
    }
}
