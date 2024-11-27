using me.fengyj.CommonLib.Utils.Scheduler;

namespace UtilsTests.Scheduler {
    [TestClass]
    public class CronExpressionTest {

        /// <summary>
        /// Test method for 'CronExpression.IsSatisfiedBy(DateTime)'.
        /// </summary>
        [TestMethod]
        public void TestIsSatisfiedBy() {
            var cronExpression = new CronExpression("0 15 10 * * ? 2005");

            var cal = new DateTime(2005, 6, 1, 10, 15, 0).ToUniversalTime();
            Assert.IsTrue(cronExpression.IsSatisfiedBy(cal));

            cal = cal.AddYears(1);
            Assert.IsFalse(cronExpression.IsSatisfiedBy(cal));

            cal = new DateTime(2005, 6, 1, 10, 16, 0).ToUniversalTime();
            Assert.IsFalse(cronExpression.IsSatisfiedBy(cal));

            cal = new DateTime(2005, 6, 1, 10, 14, 0).ToUniversalTime();
            Assert.IsFalse(cronExpression.IsSatisfiedBy(cal));

            cronExpression = new CronExpression("0 15 10 ? * MON-FRI");

            // weekends
            cal = new DateTime(2007, 6, 9, 10, 15, 0).ToUniversalTime();
            Assert.IsFalse(cronExpression.IsSatisfiedBy(cal));
            Assert.IsFalse(cronExpression.IsSatisfiedBy(cal.AddDays(1)));

            cronExpression = new CronExpression("0 0 0 1W * ? 2044,2049");
            cal = new DateTime(2044, 1, 1).ToUniversalTime();
            Assert.IsTrue(cronExpression.IsSatisfiedBy(cal));

            cronExpression = CronExpressionUtils.AtHourAndMinuteOnGivenDaysOfWeek(9, 30, DayOfWeek.Monday, DayOfWeek.Sunday);
            cal = new DateTime(2024, 11, 25, 9, 30, 0).ToUniversalTime();
            Assert.IsTrue(cronExpression.IsSatisfiedBy(cal));
            cal = cronExpression.GetNextValidTimeAfter(cal.ToUniversalTime())!.Value.LocalDateTime;
            Assert.AreEqual(DayOfWeek.Sunday, cal.DayOfWeek);
            cal = cronExpression.GetNextValidTimeAfter(cal.ToUniversalTime())!.Value.LocalDateTime;
            Assert.AreEqual(DayOfWeek.Monday, cal.DayOfWeek);
        }

        [TestMethod]
        public void TestLastDayOffset() {
            var cronExpression = new CronExpression("0 15 10 L-2 * ? 2010");

            var cal = new DateTime(2010, 10, 29, 10, 15, 0).ToUniversalTime(); // last day - 2
            Assert.IsTrue(cronExpression.IsSatisfiedBy(cal));

            cal = new DateTime(2010, 10, 28, 10, 15, 0).ToUniversalTime();
            Assert.IsFalse(cronExpression.IsSatisfiedBy(cal));

            cronExpression = new CronExpression("0 15 10 L-5W * ? 2010");

            cal = new DateTime(2010, 10, 26, 10, 15, 0).ToUniversalTime(); // last day - 5
            Assert.IsTrue(cronExpression.IsSatisfiedBy(cal));

            cronExpression = new CronExpression("0 15 10 L-1 * ? 2010");

            cal = new DateTime(2010, 10, 30, 10, 15, 0).ToUniversalTime(); // last day - 1
            Assert.IsTrue(cronExpression.IsSatisfiedBy(cal));

            cronExpression = new CronExpression("0 15 10 L-1W * ? 2010");

            cal = new DateTime(2010, 10, 29, 10, 15, 0).ToUniversalTime(); // nearest weekday to last day - 1 (29th is a friday in 2010)
            Assert.IsTrue(cronExpression.IsSatisfiedBy(cal));
        }

        [TestMethod]
        public void TestGetExpressionDescription() {

            var exp = new CronExpression("0 */5 * ? * *");

            Assert.AreEqual("Every 5 minutes", exp.GetExpressionDescription());

            exp = new CronExpression("0 15,30,45 * ? * *");

            Assert.AreEqual("At 15, 30, and 45 minutes past the hour", exp.GetExpressionDescription());

            exp = new CronExpression("0 15,30,45 * ? * FRI-SUN 2024");

            Assert.AreEqual("At 15, 30, and 45 minutes past the hour, Friday through Sunday, only in 2024", exp.GetExpressionDescription());

            exp = new CronExpression("0 0 12 ? * SAT");

            Assert.AreEqual("At 12:00, only on Saturday", exp.GetExpressionDescription());

            exp = new CronExpression("0 0 12 L-2 * ?");

            Assert.AreEqual("At 12:00, 2 days before the last day of the month", exp.GetExpressionDescription());

            exp = new CronExpression("0 0 12 ? JAN,FEB 5#3");

            Assert.AreEqual("At 12:00, on the third Friday of the month, only in January and February", exp.GetExpressionDescription());

            exp = new CronExpression("0 0 0 1W * ? 2044,2049");

            Assert.AreEqual("At 00:00, on the first weekday of the month, only in 2044 and 2049", exp.GetExpressionDescription());
        }
    }
}
