using System;
using System.Timers;


namespace tud.mci.tangram.TangramLector.Window_Manager
{
    public sealed class BlinkTimer
    {
        private int quater = 4;
        private int quater2 = 3;
        private const int interval = 100; 
        public readonly Timer timer;
        private static readonly BlinkTimer _instance = new BlinkTimer();
        public bool Set { get; private set; }

        #region Constructor / Destructor / Singleton

        BlinkTimer()
        {
            timer = new Timer();
            timer.Interval = interval;
            timer.Elapsed += timer_Elapsed;
            timer.Start();
        }

        ~BlinkTimer()
        {
            try
            {
                timer.Stop();
                timer.Dispose();
            }
            catch { }
        }

        /// <summary>
        /// Gets the singleton instance of the BlinkTimer Object.
        /// </summary>
        /// <value>The instance.</value>
        public static BlinkTimer Instance { get { return _instance; } }
        #endregion

        #region Events


        void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            quater = ++quater % 4;
            quater2 = ++quater2 % 3;
            fireQuarterTick(e);
            if ((quater % 2) == 0)
            {
                fireHalfTick(e);

                if ((quater % 4) == 0)
                {
                    Set = !Set;
                    fireTick(e);
                }
            }

            if ((quater2) == 0)
            {
                fireThreeQuarterTick(e);
            }
        }

        /// <summary>
        /// Occurs when [quarter tick].
        /// </summary>
        public event EventHandler<EventArgs> QuarterTick;
        /// <summary>
        /// Occurs when [half tick].
        /// </summary>
        public event EventHandler<EventArgs> HalfTick;
        /// <summary>
        /// Occurs when [three-quarter tick].
        /// </summary>
        public event EventHandler<EventArgs> ThreeQuarterTick;
        /// <summary>
        /// Occurs when [tick].
        /// </summary>
        public event EventHandler<EventArgs> Tick;

        private void fireQuarterTick(EventArgs e)
        {
            try
            {
                if (QuarterTick != null) QuarterTick.DynamicInvoke(this, new QuarterTickEventArgs(Set, quater));
            }
            catch { }
        }
        private void fireHalfTick(EventArgs e)
        {
            try
            {
                if (HalfTick != null) HalfTick.DynamicInvoke(this, new HalfTickEventArgs(Set, quater / 2));
            }
            catch { }
        }
        private void fireThreeQuarterTick(EventArgs e)
        {
            try
            {
                if (ThreeQuarterTick != null) ThreeQuarterTick.DynamicInvoke(this, new BlinkTickEventArgs(Set));
            }
            catch { }
        }
        private void fireTick(EventArgs e)
        {
            try
            {
                if (Tick != null) Tick.DynamicInvoke(this, new BlinkTickEventArgs(Set));
            }
            catch { }
        }
        #endregion
    }

    #region EventArg Classes

    public class BlinkTickEventArgs : EventArgs
    {
        public readonly bool Set;
        public BlinkTickEventArgs(bool set) { Set = set; }
    }

    public class QuarterTickEventArgs : BlinkTickEventArgs
    {
        public readonly int Quater;
        public QuarterTickEventArgs(bool set, int quater) : base(set) { this.Quater = quater; }
    }

    public class HalfTickEventArgs : BlinkTickEventArgs
    {
        public readonly int Half;
        public HalfTickEventArgs(bool set, int half) : base(set) { this.Half = half; }
    }
    #endregion

}