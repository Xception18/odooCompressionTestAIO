"""Thread-safe UI dispatcher for calling GUI functions from worker threads"""
import queue
from typing import Callable, Any

class UiDispatcher:
    """
    Safely dispatch UI updates from worker threads to the main Tkinter thread.
    
    Usage:
        dispatcher = UiDispatcher(root)
        
        # From worker thread:
        dispatcher.post(label.configure, text="Done")
        dispatcher.post(lambda: messagebox.showinfo("Title", "Message"))
    """
    
    def __init__(self, root, interval_ms: int = 50):
        """
        Initialize dispatcher.
        
        Args:
            root: Tk/CTk root window
            interval_ms: Pump interval in milliseconds (default: 50ms = 20 Hz)
        """
        self._root = root
        self._queue = queue.Queue()
        self._interval = interval_ms
        self._running = True
        self._pump()
    
    def post(self, fn: Callable, *args, **kwargs) -> None:
        """
        Post a function to be called on the main thread.
        
        Args:
            fn: Callable to execute on main thread
            *args: Positional arguments for fn
            **kwargs: Keyword arguments for fn
        """
        self._queue.put((fn, args, kwargs))
    
    def _pump(self):
        """Process queued calls (runs on main thread)"""
        if not self._running:
            return
        
        try:
            # Process all pending calls
            while True:
                fn, args, kwargs = self._queue.get_nowait()
                try:
                    fn(*args, **kwargs)
                except Exception as e:
                    # Log but don't crash the pump
                    print(f"[UiDispatcher] Error in queued call: {e}")
        except queue.Empty:
            pass
        
        # Schedule next pump
        self._root.after(self._interval, self._pump)
    
    def stop(self):
        """Stop the dispatcher (cleanup)"""
        self._running = False


class EventEmitter:
    """
    Simple event emitter for service → UI communication.
    
    Usage:
        emitter = EventEmitter(dispatcher)
        
        # Subscribe
        emitter.on("step_started", lambda step: print(f"Started {step}"))
        
        # Emit (from worker thread)
        emitter.emit("step_started", "Copy Data")
    """
    
    def __init__(self, dispatcher: UiDispatcher):
        self._dispatcher = dispatcher
        self._handlers = {}
    
    def on(self, event: str, handler: Callable):
        """Register event handler"""
        if event not in self._handlers:
            self._handlers[event] = []
        self._handlers[event].append(handler)
    
    def off(self, event: str, handler: Callable):
        """Unregister event handler"""
        if event in self._handlers:
            self._handlers[event].remove(handler)
    
    def emit(self, event: str, *args, **kwargs):
        """Emit event (thread-safe)"""
        if event in self._handlers:
            for handler in self._handlers[event]:
                self._dispatcher.post(handler, *args, **kwargs)
