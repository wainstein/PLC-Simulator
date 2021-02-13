using System.Drawing;
using System.Windows.Forms;

namespace PLCTools
{
    public partial class SwitchControl : UserControl
    {
        public SwitchControl()
        {
            InitializeComponent();
            this.SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            this.SetStyle(ControlStyles.DoubleBuffer, true);
            this.SetStyle(ControlStyles.ResizeRedraw, true);
            this.SetStyle(ControlStyles.Selectable, true);
            this.SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            this.SetStyle(ControlStyles.UserPaint, true);
            this.BackColor = Color.Transparent;
            this.Cursor = Cursors.Hand;
            this.Size = new Size(40, 13);
        }
        public bool isSwitch = false;
        public bool Switchable = true;

        public double quality { get; set; } = 1;
        private int Value { get; set; } = 0;

        public void EnableSwitch()
        {
            Switchable = true;
            this.Invalidate();
        }

        public void DisableSwitch()
        {
            Switchable = false;
            this.Invalidate();
        }

        public bool Pending
        {
            get
            {
                if (isSwitch)
                {
                    if (Value == 0) return true;
                    else return false;
                }
                else
                {
                    if (Value == 1) return true;
                    else return false;
                }
            }
        }
        public int value
        {
            get
            {
                if (isSwitch)
                    return 1;
                else
                    return 0;
            }
            set
            {
                if (value != 0 && value != 1)
                {
                    Value = 0;
                    fetch = true;
                    isSwitch = true;
                    return;
                }
                else
                {
                    if (Pending)
                    {
                        if (value == (isSwitch ? 0 : 1))
                        {
                            return;
                        }
                    }
                    isSwitch = value == 1 ? true : false;
                    Value = value;
                    this.Invalidate();
                }
            }
        }

        public bool fetch { get; set; } = false;


        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Rectangle rec = new Rectangle(0, 0, this.Size.Width, this.Size.Height);
            if (Enabled)
            {
                if (quality == 1)
                {
                    if (Switchable)
                    {
                        if (isSwitch)
                        {
                            if (Value == 1)
                            {
                                g.DrawImage(Properties.Resources.on, rec);
                            }
                            else
                            {
                                g.DrawImage(Properties.Resources.offton, rec);
                            }
                        }
                        else
                        {
                            if (Value == 0)
                            {
                                g.DrawImage(Properties.Resources.off, rec);
                            }
                            else
                            {
                                g.DrawImage(Properties.Resources.ontoff, rec);
                            }
                        }
                    }
                    else
                    {
                        if (isSwitch)
                        {
                            g.DrawImage(Properties.Resources.bon, rec);
                        }
                        else
                        {
                            g.DrawImage(Properties.Resources.boff, rec);
                        }
                    }
                }
                else
                {
                    g.DrawImage(Properties.Resources.bad, rec);
                }
            }
            else
            {
                g.DrawImage(Properties.Resources.no, rec);
            }
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            if (Enabled)
            {
                if (Switchable)
                {
                    if (!Pending)
                    {
                        isSwitch = !isSwitch;
                        fetch = true;
                        this.Invalidate();
                        base.OnMouseClick(e);
                    }
                }
            }
        }

        public void switchOn()
        {
            isSwitch = true;
        }
        public void switchOff()
        {
            isSwitch = false;
        }
    }
}
