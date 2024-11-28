namespace me.fengyj.CommonLib.Utils.Scheduler {
    public static class CronExpressionUtils {

        public static readonly DayOfWeek[] DefaultWeekends = { DayOfWeek.Saturday, DayOfWeek.Sunday };
        private static readonly Dictionary<DayOfWeek, string> DayOfWeeks = Enumerable.Range(0, 7)
            .Select(i => (DayOfWeek)(i))
            .ToDictionary(i => i, i => i.ToString()[..3].ToUpper());

        /// <summary>
        /// Get next valid schedule time
        /// </summary>
        /// <param name="cron"></param>
        /// <param name="from"></param>
        /// <param name="adjustOnHoliday">
        /// adjust schedule time if it's holiday. 
        /// - Null means skip it and return the next one.
        /// - 0 means don't check holiday
        /// - > 0 means return the next nth workday
        /// - < 0 means return the previous nth workday
        /// </param>
        /// <param name="dayOfWeekends"></param>
        /// <returns></returns>
        public static DateTimeOffset? GetNextValidTimeAfter(
            this CronExpression cron,
            DateTimeOffset from,
            int? adjustOnHoliday,
            DayOfWeek[]? dayOfWeekends = null) {

            if (dayOfWeekends == null) dayOfWeekends = DefaultWeekends;

            return GetNextValidTimeAfter(cron, from, (d, days) => GetBusinessDay(d, days, dayOfWeekends), adjustOnHoliday);
        }

        private static DateTimeOffset? GetBusinessDay(DateTimeOffset date, int days, DayOfWeek[] weekends) {

            var dw = date.DayOfWeek;

            if (!weekends.Contains(dw)) return date;
            else if (days == 0) return null;

            var t = Math.Abs(days);
            var step = days / t;
            var c = 0;

            for (var i = step; ; i = i + step) {
                var d = date.AddDays(i);
                dw = d.DayOfWeek;
                if (weekends.Contains(dw)) continue;
                c++;
                if (c == t)
                    return d;
            }
        }

        /// <summary>
        /// Get next valid schedule time
        /// </summary>
        /// <param name="cron"></param>
        /// <param name="from"></param>
        /// <param name="adjustOnHoliday">
        /// adjust schedule time if it's holiday. 
        /// - Null means skip it and return the next one.
        /// - 0 means don't check holiday
        /// - > 0 means return the next nth workday
        /// - < 0 means return the previous nth workday
        /// </param>
        /// <param name="dayOfWeekends"></param>
        /// <returns></returns>
        public static DateTime? GetNextValidTimeAfter(
            this CronExpression cron,
            DateTime from,
            int? adjustOnHoliday,
            DayOfWeek[]? dayOfWeekends = null) {

            if (dayOfWeekends == null) dayOfWeekends = DefaultWeekends;

            return GetNextValidTimeAfter(cron, from, (d, days) => GetBusinessDay(d, days, dayOfWeekends), adjustOnHoliday);
        }

        private static DateTime? GetBusinessDay(DateTime date, int days, DayOfWeek[] weekends) {

            var dw = date.DayOfWeek;

            if (!weekends.Contains(dw)) return date;
            else if (days == 0) return null;

            var t = Math.Abs(days);
            var step = days / t;
            var c = 0;

            for (var i = step; ; i = i + step) {
                var d = date.AddDays(i);
                dw = d.DayOfWeek;
                if (weekends.Contains(dw)) continue;
                c++;
                if (c == t)
                    return d;
            }
        }

        /// <summary>
        /// Get next valid schedule time
        /// </summary>
        /// <param name="cron">cron expression</param>
        /// <param name="from">start time</param>
        /// <param name="adjustOnHoliday">
        /// adjust schedule time if it's holiday. 
        /// - Null means skip it and return the next one.
        /// - 0 means don't check holiday
        /// - > 0 means return the next nth workday
        /// - < 0 means return the previous nth workday
        /// </param>
        /// <param name="businessDayGetter"></param>
        /// <returns></returns>
        public static DateTimeOffset? GetNextValidTimeAfter(
            this CronExpression cron,
            DateTimeOffset from,
            Func<DateTimeOffset, int, DateTimeOffset?> businessDayGetter,
            int? adjustOnHoliday) {

            DateTimeOffset? next = from;
            DateTimeOffset? businessDay = null;

            while (businessDay == null || businessDay.Value <= from) {

                next = cron.GetNextValidTimeAfter(next.Value);
                if (next == null) return null;
                if (adjustOnHoliday == 0) return next; // don't check holidays

                businessDay = businessDayGetter(next.Value, adjustOnHoliday ?? 0); // 0 means just checking if the date is business day or not. if not, return null.
                if (adjustOnHoliday == null && businessDay == next) return next; // "next" is business day
                if (adjustOnHoliday != null && adjustOnHoliday != 0 && businessDay == null) return next; // in case businessDayGetter cannot get a valid value.
            }

            return businessDay;
        }

        /// <summary>
        /// Get next valid schedule time
        /// </summary>
        /// <param name="cron">cron expression</param>
        /// <param name="from">start time</param>
        /// <param name="adjustOnHoliday">
        /// adjust schedule time if it's holiday. 
        /// - Null means skip it and return the next one.
        /// - 0 means don't check holiday
        /// - > 0 means return the next nth workday
        /// - < 0 means return the previous nth workday
        /// </param>
        /// <param name="businessDayGetter"></param>
        /// <returns></returns>
        public static DateTime? GetNextValidTimeAfter(
            this CronExpression cron,
            DateTime from,
            Func<DateTime, int, DateTime?> businessDayGetter,
            int? adjustOnHoliday) {

            DateTimeOffset? next = new DateTimeOffset(from);
            DateTime? businessDay = null;

            while (businessDay == null || businessDay.Value <= from) {

                next = cron.GetNextValidTimeAfter(next.Value);
                if (next == null) return null;
                var nextDT = from.Kind switch {
                    DateTimeKind.Utc => next.Value.UtcDateTime,
                    _ => next.Value.LocalDateTime
                };
                if (adjustOnHoliday == 0) return nextDT; // don't check holidays

                businessDay = businessDayGetter(nextDT, adjustOnHoliday ?? 0); // 0 means just checking if the date is business day or not. if not, return null.
                if (adjustOnHoliday == null && businessDay == next) return nextDT; // "next" is business day
                if (adjustOnHoliday != null && adjustOnHoliday != 0 && businessDay == null) return nextDT; // in case businessDayGetter cannot get a valid value.
            }

            return businessDay;
        }

        /// <summary>
        /// Get human readable description of the Cron expression.
        /// </summary>
        /// <param name="cron"></param>
        /// <returns></returns>
        public static string? GetExpressionDescription(this CronExpression cron) {

            return CronExpressionDescriptor.GetDescription(cron.CronExpressionString);
        }

        /// <summary>
        /// Get human readable description of the Cron expression.
        /// </summary>
        /// <param name="cron"></param>
        /// <param name="adjustOnHoliday">
        /// adjust schedule time if it's holiday. 
        /// - Null means skip it and return the next one.
        /// - 0 means don't check holiday
        /// - > 0 means return the next nth workday
        /// - < 0 means return the previous nth workday
        /// </param>
        /// <returns></returns>
        public static string? GetExpressionDescription(this CronExpression cron, int? adjustOnHoliday) {

            var desc = GetExpressionDescription(cron);
            return adjustOnHoliday switch {
                null => $"{desc} (skip holiday or weekend)",
                0 => desc,
                1 => $"{desc} (next workday if holiday or weekend)",
                > 1 => $"{desc} (next {adjustOnHoliday - 1} workdays after the first workday if holiday or weekend)",
                -1 => $"{desc} (last workday if holiday or weekend)",
                < -1 => $"{desc} (previous {Math.Abs(adjustOnHoliday.Value) - 1} workdays before the last workday if holiday or weekend)",
            };
        }

        public static CronExpression DailyAtHourAndMinute(int hour, int minute) {
            return new CronExpression($"0 {minute} {hour} ? * *");
        }

        public static CronExpression AtHourAndMinuteOnGivenDaysOfWeek(int hour, int minute, params DayOfWeek[] daysOfWeek) {
            return new CronExpression($"0 {minute} {hour} ? * {string.Join(",", daysOfWeek.OrderBy(i => (int)i).Select(i => DayOfWeeks[i]).ToArray())}");
        }

        public static CronExpression MonthlyOnDayAndHourAndMinute(int dayOfMonth, int hour, int minute) {

            if (dayOfMonth > 0)
                return new CronExpression($"0 {minute} {hour} {dayOfMonth} * ?");
            else if (dayOfMonth == -1)
                return new CronExpression($"0 {minute} {hour} L * ?");
            else if (dayOfMonth < -1)
                return new CronExpression($"0 {minute} {hour} L{dayOfMonth + 1} * ?");
            else
                throw new ArgumentException("Value cannot be 0", nameof(dayOfMonth));
        }

        public static CronExpression QuarterlyOnDayAndHourAndMinute(int monthOfStart, int dayOfMonth, int hour, int minute) {

            monthOfStart = monthOfStart > 3 ? monthOfStart % 3 : monthOfStart;
            if (dayOfMonth > 0)
                return new CronExpression($"0 {minute} {hour} {dayOfMonth} {monthOfStart},{monthOfStart + 3},{monthOfStart + 6},{monthOfStart + 9} ?");
            else if (dayOfMonth == -1)
                return new CronExpression($"0 {minute} {hour} L {monthOfStart},{monthOfStart + 3},{monthOfStart + 6},{monthOfStart + 9} ?");
            else if (dayOfMonth < -1)
                return new CronExpression($"0 {minute} {hour} L{dayOfMonth + 1} {monthOfStart},{monthOfStart + 3},{monthOfStart + 6},{monthOfStart + 9} ?");
            else
                throw new ArgumentException("Value cannot be 0", nameof(dayOfMonth));
        }

        public static CronExpression AnnualyOnMonthAndDayAndHourAndMinute(int month, int dayOfMonth, int hour, int minute) {

            if (dayOfMonth > 0)
                return new CronExpression($"0 {minute} {hour} {dayOfMonth} {month} ?");
            else if (dayOfMonth == -1)
                return new CronExpression($"0 {minute} {hour} L {month} ?");
            else if (dayOfMonth < -1)
                return new CronExpression($"0 {minute} {hour} L{dayOfMonth + 1} {month} ?");
            else
                throw new ArgumentException("Value cannot be 0", nameof(dayOfMonth));
        }

    }
}
