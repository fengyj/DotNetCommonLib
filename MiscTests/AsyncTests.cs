namespace MiscTests {
    [TestClass]
    public class AsyncTests {

        private Random random = new(5);

        [TestMethod]
        public void Test_Async() {
            this.RunJobsSeriallyAsync().GetAwaiter().GetResult();
            this.RunJobsParallelAsync().GetAwaiter().GetResult();
            this.RunJobsInTwoThreadsAsync().GetAwaiter().GetResult();
        }

        private async Task RunJobsSeriallyAsync() {

            var cancelToken = new CancellationToken();
            var t = DateTime.Now;
            Console.WriteLine("Run jobs in serial");
            foreach (var job in Enumerable.Range(0, 10)) {
                await this.DoJob(job, cancelToken);
            }
            Console.WriteLine($"Done, cost {(int)(DateTime.Now - t).TotalSeconds} seconds");
        }

        private async Task RunJobsParallelAsync() {

            var cancelToken = new CancellationToken();
            var t = DateTime.Now;
            Console.WriteLine("Run jobs in parallel");
            var jobs = Enumerable.Range(0, 10).Select(i => this.DoJob(i, cancelToken)).ToArray();
            await Task.WhenAll(jobs);
            Console.WriteLine($"Done, cost {(int)(DateTime.Now - t).TotalSeconds} seconds");
        }

        private async Task RunJobsInTwoThreadsAsync() {
            var t = DateTime.Now;
            Console.WriteLine("Run jobs in 2 threads");
            await Parallel.ForEachAsync(
                Enumerable.Range(0, 10),
                new ParallelOptions { MaxDegreeOfParallelism = 2 },
                async (job, token) => await this.DoJob(job, token));
            Console.WriteLine($"Done, cost {(int)(DateTime.Now - t).TotalSeconds} seconds");
        }

        private async Task<int> DoJob(int jobId, CancellationToken token) {
            var t = DateTime.Now;
            Console.WriteLine($"Job {jobId,2} - started at {t:mm:ss.fff}");
            await Task.Delay(TimeSpan.FromSeconds(this.random.NextDouble() * 4.5 + 0.5), token);
            var result = (int)(DateTime.Now - t).TotalSeconds;
            Console.WriteLine($"Job {jobId,2} - end at {DateTime.Now:mm:ss.fff}, cost {result} second");
            return result;
        }
    }
}