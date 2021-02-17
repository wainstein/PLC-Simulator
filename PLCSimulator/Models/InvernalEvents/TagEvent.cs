using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PLCTools.Service.InvernalEvents
{
    public class TagEvent
    {
        public string Tagevent { get; set; }
        public Boolean Activated { get; set; }
        public Boolean Deactivated { get; set; }
        public List<List<TablesCommand>> Targets { get; set; }
        public List<List<string>> ScriptsList { get; set; }
        public bool Enable { get; set; }
    }
}
