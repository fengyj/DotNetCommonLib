using me.fengyj.CommonLib.Utils.Scheduler;

namespace UtilsTests.Scheduler {
    [TestClass]
    public class CronExpressionUtilsTests {

        [TestMethod]
        public void TestGetNextValidTimeAfter_DateTimeOffset() {

            var cron = CronExpressionUtils.DailyAtHourAndMinute(10, 0);
            // 2024-11-22 is Friday
            var next = cron.GetNextValidTimeAfter(new DateTimeOffset(2024, 11, 22, 11, 0, 0, cron.TimeZone.BaseUtcOffset), null);

            Assert.AreEqual(new DateTimeOffset(2024, 11, 25, 10, 0, 0, cron.TimeZone.BaseUtcOffset), next);

            next = cron.GetNextValidTimeAfter(new DateTimeOffset(2024, 11, 22, 10, 0, 0, cron.TimeZone.BaseUtcOffset), 0);

            Assert.AreEqual(new DateTimeOffset(2024, 11, 23, 10, 0, 0, cron.TimeZone.BaseUtcOffset), next);

            next = cron.GetNextValidTimeAfter(new DateTimeOffset(2024, 11, 22, 10, 0, 0, cron.TimeZone.BaseUtcOffset), 2);

            Assert.AreEqual(new DateTimeOffset(2024, 11, 26, 10, 0, 0, cron.TimeZone.BaseUtcOffset), next);

            next = cron.GetNextValidTimeAfter(new DateTimeOffset(2024, 11, 22, 10, 0, 0, cron.TimeZone.BaseUtcOffset), -1);

            Assert.AreEqual(new DateTimeOffset(2024, 11, 25, 10, 0, 0, cron.TimeZone.BaseUtcOffset), next);

            next = cron.GetNextValidTimeAfter(new DateTimeOffset(2024, 11, 21, 10, 0, 0, cron.TimeZone.BaseUtcOffset), -1);

            Assert.AreEqual(new DateTimeOffset(2024, 11, 22, 10, 0, 0, cron.TimeZone.BaseUtcOffset), next);

            next = cron.GetNextValidTimeAfter(new DateTimeOffset(2024, 11, 22, 9, 0, 0, cron.TimeZone.BaseUtcOffset), -1);

            Assert.AreEqual(new DateTimeOffset(2024, 11, 22, 10, 0, 0, cron.TimeZone.BaseUtcOffset), next);

            cron = CronExpressionUtils.MonthlyOnDayAndHourAndMinute(-1, 10, 0);

            next = cron.GetNextValidTimeAfter(new DateTimeOffset(2024, 11, 22, 9, 0, 0, cron.TimeZone.BaseUtcOffset), -1);
            // 2024-11-30 is weekend
            Assert.AreEqual(new DateTimeOffset(2024, 11, 29, 10, 0, 0, cron.TimeZone.BaseUtcOffset), next);

            next = cron.GetNextValidTimeAfter(new DateTimeOffset(2024, 11, 29, 10, 0, 0, cron.TimeZone.BaseUtcOffset), -1);

            Assert.AreEqual(new DateTimeOffset(2024, 12, 31, 10, 0, 0, cron.TimeZone.BaseUtcOffset), next);
        }

        [TestMethod]
        public void TestGetNextValidTimeAfter_DateTime() {

            var cron = CronExpressionUtils.DailyAtHourAndMinute(10, 0);
            // 2024-11-22 is Friday
            var next = cron.GetNextValidTimeAfter(new DateTime(2024, 11, 22, 11, 0, 0), null);

            Assert.AreEqual(new DateTime(2024, 11, 25, 10, 0, 0), next);

            next = cron.GetNextValidTimeAfter(new DateTime(2024, 11, 22, 10, 0, 0), 0);

            Assert.AreEqual(new DateTime(2024, 11, 23, 10, 0, 0), next);

            next = cron.GetNextValidTimeAfter(new DateTime(2024, 11, 22, 10, 0, 0), 2);

            Assert.AreEqual(new DateTime(2024, 11, 26, 10, 0, 0), next);

            next = cron.GetNextValidTimeAfter(new DateTime(2024, 11, 22, 10, 0, 0), -1);

            Assert.AreEqual(new DateTime(2024, 11, 25, 10, 0, 0), next);

            next = cron.GetNextValidTimeAfter(new DateTime(2024, 11, 21, 10, 0, 0), -1);

            Assert.AreEqual(new DateTime(2024, 11, 22, 10, 0, 0), next);

            next = cron.GetNextValidTimeAfter(new DateTime(2024, 11, 22, 9, 0, 0), -1);

            Assert.AreEqual(new DateTime(2024, 11, 22, 10, 0, 0), next);

            cron = CronExpressionUtils.MonthlyOnDayAndHourAndMinute(-1, 10, 0);

            next = cron.GetNextValidTimeAfter(new DateTime(2024, 11, 22, 9, 0, 0), -1);
            // 2024-11-30 is weekend
            Assert.AreEqual(new DateTime(2024, 11, 29, 10, 0, 0), next);

            next = cron.GetNextValidTimeAfter(new DateTime(2024, 11, 29, 10, 0, 0), -1);

            Assert.AreEqual(new DateTime(2024, 12, 31, 10, 0, 0), next);
        }
    }
}
