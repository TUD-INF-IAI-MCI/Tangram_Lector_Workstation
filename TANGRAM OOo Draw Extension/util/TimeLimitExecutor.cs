using System;
using System.Threading;
using System.Diagnostics;
using System.Threading.Tasks;

namespace tud.mci.tangram.util
{
    public static class TimeLimitExecutor
    {
        public const string PRE_THREAD_IDENTIFIER = "TLE_";

        /// <summary>
        /// Executes an <code>Action</code> with a time limit of 2000 milliseconds in an seperate asynchronous Thread.
        /// </summary>
        /// <param name="codeBlock">The code block. Us a e.g. an annoymous delegate such as
        /// <code>() =&gt; { /* Write your time bounded code here*/ }</code></param>
        /// <param name="name">The name.</param>
        /// <returns>the executing Thread</returns>
        public static Thread ExecuteWithTimeLimit(Action codeBlock, string name = "") { return ExecuteWithTimeLimit(2000, codeBlock, name); }

        /// <summary>
        /// Executes an <code>Action</code> with a time limit in an seperate asynchronous Thread.
        /// </summary>
        /// <param name="maxTime">The maximum execution time in milliseconds. After this time the thread will be abort!</param>
        /// <param name="codeBlock">The code block. Us a e.g. an annoymous delegate such as
        /// <code>() =&gt; { /* Write your time bounded code here*/ }</code></param>
        /// <param name="name">The name.</param>
        /// <returns>The executing Thread</returns>
        public static Thread ExecuteWithTimeLimit(int maxTime, Action codeBlock, string name = "")
        {
            return innerExecution(maxTime, codeBlock, name);
        }


        /// <summary>
        /// Executes an <code>Action</code> with a time limit in an seperate asynchronous Thread and waits until it ends for the return.
        /// </summary>
        /// <param name="maxTime">The maximum execution time in milliseconds. After this time the thread will be abort!</param>
        /// <param name="codeBlock">The code block. Us a e.g. an annoymous delegate such as
        /// <code>() =&gt; { /* Write your time bounded code here*/ }</code></param>
        /// <param name="name">The name.</param>
        /// <returns>indication if the execution was done successfully</returns>
        public static bool WaitForExecuteWithTimeLimit(int maxTime, Action codeBlock, string name = "")
        {
            bool success = false;
            try
            {
                var executor = TimeLimitExecutor.ExecuteWithTimeLimit(maxTime, codeBlock, name);

                while (executor != null && executor.IsAlive && executor.ThreadState == System.Threading.ThreadState.Running)
                {
                    Thread.Sleep(1);
                }
                success = true;
            }
            catch (System.Threading.ThreadAbortException)
            {
                return false;
            }

            return success;
        }


        /// <summary>
        /// starts the execution and the observation of the maximum runtime.
        /// </summary>
        /// <param name="timeSpan">The time span.</param>
        /// <param name="codeBlock">The code block.</param>
        private static Thread innerExecution(int timeSpan, Action codeBlock, string name = "")
        {

//#if LIBRE

//            codeBlock.Invoke();
//            return null;
//#else

            if (Thread.CurrentThread != null && Thread.CurrentThread.Name != null && Thread.CurrentThread.Name.StartsWith(PRE_THREAD_IDENTIFIER))
            {
                codeBlock.Invoke();
                return null;
            }
            else
            {

                int sleep = Math.Max(10, (timeSpan / 100));
                Thread task = new Thread(new ThreadStart(codeBlock));
                task.Name = PRE_THREAD_IDENTIFIER + name;
                //task.Priority = ThreadPriority.AboveNormal;
                task.Start();

                // start the observation
                try
                {
                    useStopwatch();
                    long startTime = getStartTime();
                    Thread thread = new Thread(delegate()
                    {
                        while (true)
                        {
                            if (task == null
                                || task.ThreadState == System.Threading.ThreadState.Aborted
                                || task.ThreadState == System.Threading.ThreadState.Stopped
                                || task.ThreadState == System.Threading.ThreadState.Unstarted)
                            {
                                // if execution has already ended
                                return;
                            }
                            if (timeIsElapsed(startTime, timeSpan))
                            {
                                // kill the hanging Thread
                                freeStopwatch();
                                task.Interrupt();
                                task.Abort();
                                Logger.Instance.Log(LogPriority.DEBUG, "TimeLimitExecutor", "[ERROR] Have to kill the execution of task '" + name + "'");
                                return;
                            }
                            else
                            {
                                Thread.Sleep(sleep);
                            }
                        }
                    });

                    // start the observation
                    thread.Start();
                    //sw.Start();
                }
                catch (System.Threading.ThreadAbortException) { }
                return task;
            } 
//#endif
        }

        /// <summary>
        /// get the start time of the current running stopwatch
        /// </summary>
        /// <returns>elapsed milliseconds since the last start of the stopwatch</returns>
        private static long getStartTime(){
            long startTime = 0;

            if (_stopwatch.IsRunning) startTime = _stopwatch.ElapsedMilliseconds + 1;
            else { _stopwatch.Start(); }

            return startTime;
        }


        static volatile Stopwatch _stopwatch = new Stopwatch();
        /// <summary>
        /// Uses the stopwatch. Sets a reference counter so the stopwatch can bee
        /// </summary>
        /// <returns>the global stopwatch</returns>
        private static Stopwatch useStopwatch()
        {
            _refCounter++;
            return _stopwatch;
        }

        private static volatile int _refCounter = 0;
        /// <summary>
        /// releases the stopwatch (decrement the reference counter) and reset the 
        /// stopwatch if there is no referencing task using it.
        /// </summary>
        private static void freeStopwatch(){
            _refCounter--;
            if (_refCounter < 1) _stopwatch.Stop();
        }


        /// <summary>
        /// Determines if the duration time is elapsed or not
        /// </summary>
        /// <param name="startTime">The start time.</param>
        /// <param name="duration">The amximal duration.</param>
        /// <returns></returns>
        private static bool timeIsElapsed(long startTime, int duration)
        {
            return (_stopwatch.ElapsedMilliseconds - startTime > duration);
        }
    }
}
