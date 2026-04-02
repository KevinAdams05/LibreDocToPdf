/**
 * Simple async semaphore for limiting concurrency in GLib async methods.
 */
public class AsyncSemaphore {
    private int max_count;
    private int current_count;
    private Queue<SourceFunc> waiters;

    public AsyncSemaphore (int max_count) {
        this.max_count = max_count;
        this.current_count = 0;
        this.waiters = new Queue<SourceFunc> ();
    }

    public async void acquire () {
        if (current_count < max_count) {
            current_count++;
            return;
        }

        SourceFunc callback = acquire.callback;
        waiters.push_tail ((owned) callback);
        yield;
    }

    public void release () {
        current_count--;

        if (!waiters.is_empty ()) {
            current_count++;
            var callback = waiters.pop_head ();
            Idle.add ((owned) callback);
        }
    }
}

/**
 * Simple async task completion source for waiting on async operations.
 */
public class AsyncTask {
    private bool completed = false;
    private SourceFunc? waiter = null;

    public async void wait_async () {
        if (completed) return;

        waiter = wait_async.callback;
        yield;
    }

    public void complete () {
        completed = true;

        if (waiter != null) {
            var cb = (owned) waiter;
            waiter = null;
            Idle.add ((owned) cb);
        }
    }
}
