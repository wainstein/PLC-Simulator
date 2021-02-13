using System;

namespace PLCTools.Models
{
    public class InternalTag
    {
        public string Tag { get; set; }
        public object Value { get; set; }
        public Boolean Activated { get; set; }
        public Boolean Deactivated { get; set; }
        public object LastValue { get; set; }
        public Boolean IsIntTag { get; set; }
    }
}
