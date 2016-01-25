using System;
using System.Drawing;
using System.Windows.Forms;

namespace tud.mci.tangram.TangramLector
{
    public partial class TransparentOverlayPictureBox : PictureBox
    {
        private static readonly TransparentOverlayPictureBox _instance = new TransparentOverlayPictureBox();
        public static TransparentOverlayPictureBox Instance { get { return _instance; } }

        public TransparentOverlayPictureBox()
        {
            InitializeComponent();
        }

        // always handle mouse interaction transparent
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0084/*WM_NCHITTEST*/)
                m.Result = (IntPtr)(-1)/*HTTRANSPARENT*/;
            else base.WndProc(ref m);
        }

        // also paint bounding box
        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
            // get the graphics object to use to draw
            Graphics g = pe.Graphics;
            g.DrawRectangle(new Pen(Color.DarkRed, 1.0f), new Rectangle(Location,new Size(Width-1,Height-1)));
        }
    }
}
