using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace tud.mci.tangram.models.Interfaces
{
    interface IUpdateable
    {
        /// <summary>
        /// Updates this instance and his related Objects.
        /// </summary>
        void Update();
    }

    interface IDisposingObserver
    {
        /// <summary>
        /// Gets a value indicating whether this instance is disposed.
        /// </summary>
        /// <value><c>true</c> if disposed; otherwise, <c>false</c>.</value>
        bool Disposed { get; }
        /// <summary>
        /// Occurs when this object is disposing.
        /// </summary>
        event EventHandler ObserverDisposing;
    }

}