using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelAddInTest.Utils
{
    public static class TaskExtensions
    {
        // Await a Task<T> with a CancellationToken on .NET Framework
        public static async Task<T> WithCancellation<T>(this Task<T> task, CancellationToken ct)
        {
            if (task == null) throw new ArgumentNullException(nameof(task));
            if (task.IsCompleted) return await task; // fast-path

            var tcs = new TaskCompletionSource<bool>();
            using (ct.Register(s => ((TaskCompletionSource<bool>)s).TrySetResult(true), tcs))
            {
                if (task != await Task.WhenAny(task, tcs.Task))
                    throw new OperationCanceledException(ct);
            }
            return await task; // already completed
        }
    }
}
