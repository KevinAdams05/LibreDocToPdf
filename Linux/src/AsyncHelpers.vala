using GLib;
using Gee;

/**
 * GTK4-friendly async semaphore using Futures-style waiting.
 */
public class AsyncSemaphore : Object {
    private int max_count;
    private int current_count = 0;
    private Gee.Queue<TaskCompletionSource> waiters;

    public AsyncSemaphore (int max_count) {
        this.max_count = max_count;
        this.waiters = new ArrayQueue<TaskCompletionSource> ();
    }

    public async void acquire () {
        // Fast path
        if (current_count < max_count) {
            current_count++;
            return;
        }

        // Create a waiter and suspend
        var tcs = new TaskCompletionSource ();
        waiters.offer (tcs);

        yield tcs.wait_async ();
    }

    public void release () {
        if (!waiters.is_empty) {
            // Wake next waiter instead of freeing slot
            var tcs = waiters.poll ();

            Idle.add (() => {
                tcs.complete ();
                return false;
            });
        } else {
            // No waiters → free slot
            if (current_count > 0)
                current_count--;
        }
    }
}

/**
 * GTK4-style async completion primitive.
 * (Cleaner replacement for manual SourceFunc handling)
 */
public class TaskCompletionSource : Object {
    private bool completed = false;
    private SourceFunc? continuation = null;

    public async void wait_async () {
        if (completed)
            return;

        continuation = wait_async.callback;
        yield;
    }

    public void complete () {
        if (completed)
            return;

        completed = true;

        if (continuation != null) {
            var cb = continuation;
            continuation = null;

            Idle.add (() => {
                cb ();
                return false;
            });
        }
    }
}
