using System;

namespace PLCTools.Service
{
    public class OPCItems
    {
        public string PLCName { get; set; }
        //public long Error { get; set; }
        public string Address { get; set; }
        public object Value { get; set; }
        public int Quality { get; set; }
        public string Description { get; set; }
        private bool ValuPrior { get; set; } = false;
        public string Tag { get; set; }
        public bool Activated()
        {
            if (Convert.ToInt32(Value) == 1 && !ValuPrior)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool Deactivated()
        {
            if (Convert.ToInt32(Value) != 1 && ValuPrior)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
