namespace me.fengyj.CommonLib.Utils {

    public class RetryUtil {

        public static void Execute<E>(Action action, RetryPolicy policy, Action<E, int> exceptionHandler)
            where E : Exception {

            for (var t = 1; t <= policy.MaxRetryTimes; t++) {
                try {
                    action();
                    return;
                }
                catch (E ex) {

                    if (exceptionHandler != null) exceptionHandler(ex, t);

                    if (t < policy.MaxRetryTimes)
                        Thread.Sleep(policy.Interval);
                    else
                        throw;
                }
            }

            throw new ApplicationException(); // just for fixing the compile error 
        }

        public static void Execute<E1, E2>(Action action, RetryPolicy policy, Action<E1, int> exceptionHandlerForE1, Action<E2, int> exceptionHandlerForE2)
            where E1 : Exception
            where E2 : Exception {

            CheckExceptionTypes(typeof(E1), typeof(E2));

            for (var t = 1; t <= policy.MaxRetryTimes; t++) {
                try {
                    action();
                    return;
                }
                catch (E1 ex) {

                    if (exceptionHandlerForE1 != null) exceptionHandlerForE1(ex, t);

                    if (t < policy.MaxRetryTimes)
                        Thread.Sleep(policy.Interval);
                    else
                        throw;
                }
                catch (E2 ex) {

                    if (exceptionHandlerForE2 != null) exceptionHandlerForE2(ex, t);

                    if (t < policy.MaxRetryTimes)
                        Thread.Sleep(policy.Interval);
                    else
                        throw;
                }
            }

            throw new ApplicationException(); // just for fixing the compile error 
        }

        public static T Execute<T, E>(Func<T> func, RetryPolicy policy, Action<E, int> exceptionHandler)
            where E : Exception {

            for (var t = 1; t <= policy.MaxRetryTimes; t++) {
                try {
                    return func();
                }
                catch (E ex) {

                    if (exceptionHandler != null) exceptionHandler(ex, t);

                    if (t < policy.MaxRetryTimes)
                        Thread.Sleep(policy.Interval);
                    else
                        throw;
                }
            }

            throw new ApplicationException(); // just for fixing the compile error 
        }

        public static T Execute<T, E1, E2>(Func<T> func, RetryPolicy policy, Action<E1, int> exceptionHandlerForE1, Action<E2, int> exceptionHandlerForE2)
            where E1 : Exception
            where E2 : Exception {

            CheckExceptionTypes(typeof(E1), typeof(E2));

            for (var t = 1; t <= policy.MaxRetryTimes; t++) {
                try {
                    return func();
                }
                catch (E1 ex) {

                    if (exceptionHandlerForE1 != null) exceptionHandlerForE1(ex, t);

                    if (t < policy.MaxRetryTimes)
                        Thread.Sleep(policy.Interval);
                    else
                        throw;
                }
                catch (E2 ex) {

                    if (exceptionHandlerForE2 != null) exceptionHandlerForE2(ex, t);

                    if (t < policy.MaxRetryTimes)
                        Thread.Sleep(policy.Interval);
                    else
                        throw;
                }
            }

            throw new ApplicationException(); // just for fixing the compile error 
        }

        public static async Task ExecuteAsync<E>(Func<Task> task, RetryPolicy policy, Action<E, int> exceptionHandler)
            where E : Exception {

            for (var t = 1; t <= policy.MaxRetryTimes; t++) {
                try {
                    await task();
                    return;
                }
                catch (E ex) {

                    if (exceptionHandler != null) exceptionHandler(ex, t);

                    if (t < policy.MaxRetryTimes)
                        await Task.Delay(policy.Interval);
                    else
                        throw;
                }
            }

            throw new ApplicationException(); // just for fixing the compile error 
        }

        public static async Task ExecuteAsync<E1, E2>(Func<Task> task, RetryPolicy policy, Action<E1, int> exceptionHandlerForE1, Action<E2, int> exceptionHandlerForE2)
            where E1 : Exception
            where E2 : Exception {

            CheckExceptionTypes(typeof(E1), typeof(E2));

            for (var t = 1; t <= policy.MaxRetryTimes; t++) {
                try {
                    await task();
                    return;
                }
                catch (E1 ex) {

                    if (exceptionHandlerForE1 != null) exceptionHandlerForE1(ex, t);

                    if (t < policy.MaxRetryTimes)
                        await Task.Delay(policy.Interval);
                    else
                        throw;
                }
                catch (E2 ex) {

                    if (exceptionHandlerForE2 != null) exceptionHandlerForE2(ex, t);

                    if (t < policy.MaxRetryTimes)
                        await Task.Delay(policy.Interval);
                    else
                        throw;
                }
            }

            throw new ApplicationException(); // just for fixing the compile error 
        }

        public static async Task<T> ExecuteAsync<T, E>(Func<Task<T>> task, RetryPolicy policy, Action<E, int> exceptionHandler)
            where E : Exception {

            for (var t = 1; t <= policy.MaxRetryTimes; t++) {
                try {
                    return await task();
                }
                catch (E ex) {

                    if (exceptionHandler != null) exceptionHandler(ex, t);

                    if (t < policy.MaxRetryTimes)
                        await Task.Delay(policy.Interval);
                    else
                        throw;
                }
            }

            throw new ApplicationException(); // just for fixing the compile error 
        }

        public static async Task<T> ExecuteAsync<T, E1, E2>(Func<Task<T>> task, RetryPolicy policy, Action<E1, int> exceptionHandlerForE1, Action<E2, int> exceptionHandlerForE2)
            where E1 : Exception
            where E2 : Exception {

            CheckExceptionTypes(typeof(E1), typeof(E2));

            for (var t = 1; t <= policy.MaxRetryTimes; t++) {
                try {
                    return await task();
                }
                catch (E1 ex) {

                    if (exceptionHandlerForE1 != null) exceptionHandlerForE1(ex, t);

                    if (t < policy.MaxRetryTimes)
                        await Task.Delay(policy.Interval);
                    else
                        throw;
                }
                catch (E2 ex) {

                    if (exceptionHandlerForE2 != null) exceptionHandlerForE2(ex, t);

                    if (t < policy.MaxRetryTimes)
                        await Task.Delay(policy.Interval);
                    else
                        throw;
                }
            }

            throw new ApplicationException(); // just for fixing the compile error 
        }

        private static void CheckExceptionTypes(params Type[] exceptionTypes) {

            CheckExceptionTypes(exceptionTypes, 0);
        }

        private static void CheckExceptionTypes(Type[] exceptionTypes, int startCheckingPos) {

            var T1 = exceptionTypes[startCheckingPos];

            for (var posOfT2 = startCheckingPos + 1; posOfT2 < exceptionTypes.Length; posOfT2++) {

                var T2 = exceptionTypes[posOfT2];

                if (T2.IsSubclassOf(T1))
                    throw new ArgumentException($"{T2.Name} is subclass of {T1.Name}, so it needs to be placed ahead.");
            }
        }
    }

    public class RetryPolicy {

        public static RetryPolicy Default = new RetryPolicy { MaxRetryTimes = 3, Interval = TimeSpan.FromSeconds(3) };

        public int MaxRetryTimes { get; set; }
        public TimeSpan Interval { get; set; }
    }
}