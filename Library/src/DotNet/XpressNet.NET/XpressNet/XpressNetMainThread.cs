//
//  XpressNetMainThread.cs
//  XpressNet.NET
//
//  Created by Aiko Pras on 01-11-2026.
//  Copyright (c) 2026 by Aiko Pras. All rights reserved.
//
//  The thread model of the library, and the timer class (as XpressNet-Class-Timer.cls of the VB6 DLL).
//
//  In the VB6 DLL (and the macOS version) everything runs on a single thread: the main thread (UI thread)
//  of the application. The .NET version follows the same model:
//  - All code of the library runs on the main thread: the thread that created the XpressNetClass object
//    (for example in Form_Load).
//  - Background threads do only one thing: they receive bytes (TCP, UDP, serial) or timer ticks, and pass
//    these via Post to the main thread. Post is the only transition between threads.
//  - Events are therefore raised on the main thread. The application may update its user interface
//    directly from an event handler, without Invoke.
//  - If the application calls the library from another thread, the call is passed to the main thread
//    as well, and a line is written in the logfile (see XpressNetClass.otherThread).
//

using System;
using System.Threading;
using System.Windows.Forms;

namespace XpressNet
{
  // ****************************************************************************************************************
  // ************************************************* MAIN THREAD **************************************************
  // ****************************************************************************************************************
  internal class MainThreadClass
  {
    private readonly SynchronizationContext _context;   // The context of the main thread
    private readonly int _threadId;                     // The id of the main thread
    internal volatile bool Closed;                      // After Dispose, posted actions are ignored

    internal MainThreadClass()
    {
      // In a Windows Forms application the context of the UI thread exists as soon as the first control (form)
      // has been created. If XpressNetClass is created before that moment, we install that context ourselves.
      if (SynchronizationContext.Current == null)
        SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
      _context = SynchronizationContext.Current;
      _threadId = Thread.CurrentThread.ManagedThreadId;
    }

    // Do we run on the main thread?
    internal bool IsMainThread
    {
      get { return Thread.CurrentThread.ManagedThreadId == _threadId; }
    }

    // Pass an action to the main thread. May be called from every thread
    internal void Post(Action action)
    {
      _context.Post(state => { if (!Closed) action(); }, null);
    }

    // As Post, but also after Dispose (to pass an error of the application back to the application)
    internal void PostAlways(Action action)
    {
      _context.Post(state => action(), null);
    }
  }


  // ****************************************************************************************************************
  // **************************************************** TIMER *****************************************************
  // ****************************************************************************************************************
  // Same behaviour as the XTimerClass of the VB6 DLL:
  // - The timer is periodic: it fires every Interval msec, as long as it is Enabled
  // - Setting Interval or Enabled restarts the timer. The timer runs only if Enabled is true and Interval > 0
  // - The Timer event is raised on the main thread
  // A tick that was already underway while the timer was stopped or restarted is ignored (_generation)
  internal class XTimerClass
  {
    internal event Action Timer;                        // Raised on the main thread

    private readonly MainThreadClass _mainThread;
    private System.Threading.Timer _timer;              // Fires on a thread of the thread pool
    private bool _enabled = true;                       // As in VB6: initially enabled, but without interval
    private int _interval = 0;
    private int _generation = 0;                        // Incremented at every (re)start or stop

    internal XTimerClass(MainThreadClass mainThread)
    {
      _mainThread = mainThread;
    }

    internal bool Enabled
    {
      get { return _enabled; }
      set { _enabled = value; restart(); }
    }

    internal int Interval
    {
      get { return _interval; }
      set { if (value >= 0) { _interval = value; restart(); } }
    }

    // Stops the timer for good (Dispose)
    internal void close()
    {
      _enabled = false;
      restart();
    }

    private void restart()
    {
      _generation++;
      if (_timer != null) {
        _timer.Dispose();
        _timer = null;
      }
      if (_enabled && (_interval > 0)) {
        int generation = _generation;
        _timer = new System.Threading.Timer(state => _mainThread.Post(() => tick(generation)), null, _interval, _interval);
      }
    }

    private void tick(int generation)
    {
      if (generation != _generation) return;            // the timer has been stopped or restarted in the meantime
      Action handler = Timer;
      if (handler != null) handler();
    }
  }
}
