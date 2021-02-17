using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PLCTools.Service.InvernalEvents
{
    public class ColumnsCommand
    {
        public string name { get; set; }
        public string Formula { get; set; }
        public string Format { get; set; }
        public Boolean Enabled { get; set; }
        public string KeepInTag { get; set; }         //***JB KeepInTag
        public string WriteToTag { get; set; }
        public string TagAddress { get; set; }
        public string ServerName { get; set; }
        public string PLCName { get; set; }
        public int LastValue { get; set; }             //***FL Write to OPC
        public string Value { get; set; }
    }
}
