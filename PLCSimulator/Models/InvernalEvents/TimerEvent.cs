using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PLCTools.Service.InvernalEvents
{
    public class TimerEvent
    {
        public string Tagevent { get; set; }
        public int TimerPeriod { get; set; }
        public Boolean Activated { get; set; }
        public Boolean Deactivated { get; set; }
        public DateTime LastTimeActivated { get; set; }
        public List<List<TablesCommand>> Targets { get; set; }
        public List<List<string>> ScriptsList { get; set; }
        public bool enable;
    }
}
