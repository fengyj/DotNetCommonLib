using System.Collections;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;

namespace me.fengyj.CommonLib.Utils.Scheduler {

    // COPIED FROM https://github.com/quartznet/quartznet/blob/main/src/Quartz/CronExpression.cs


    /// <summary>
    /// Provides a parser and evaluator for unix-like cron expressions. Cron
    /// expressions provide the ability to specify complex time combinations such as
    /// &quot;At 8:00am every Monday through Friday&quot; or &quot;At 1:30am every
    /// last Friday of the month&quot;.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Cron expressions are comprised of 6 required fields and one optional field
    /// separated by white space. The fields respectively are described as follows:
    /// </para>
    /// <table cellspacing="8">
    /// <tr>
    /// <th align="left">Field Name</th>
    /// <th align="left"> </th>
    /// <th align="left">Allowed Values</th>
    /// <th align="left"> </th>
    /// <th align="left">Allowed Special Characters</th>
    /// </tr>
    /// <tr>
    /// <td align="left">Seconds</td>
    /// <td align="left"> </td>
    /// <td align="left">0-59</td>
    /// <td align="left"> </td>
    /// <td align="left">, - /// /</td>
    /// </tr>
    /// <tr>
    /// <td align="left">Minutes</td>
    /// <td align="left"> </td>
    /// <td align="left">0-59</td>
    /// <td align="left"> </td>
    /// <td align="left">, - /// /</td>
    /// </tr>
    /// <tr>
    /// <td align="left">Hours</td>
    /// <td align="left"> </td>
    /// <td align="left">0-23</td>
    /// <td align="left"> </td>
    /// <td align="left">, - /// /</td>
    /// </tr>
    /// <tr>
    /// <td align="left">Day-of-month</td>
    /// <td align="left"> </td>
    /// <td align="left">1-31</td>
    /// <td align="left"> </td>
    /// <td align="left">, - /// ? / L W C</td>
    /// </tr>
    /// <tr>
    /// <td align="left">Month</td>
    /// <td align="left"> </td>
    /// <td align="left">1-12 or JAN-DEC</td>
    /// <td align="left"></td>
    /// <td align="left">, - /// /</td>
    /// </tr>
    /// <tr>
    /// <td align="left">Day-of-Week</td>
    /// <td align="left"> </td>
    /// <td align="left">1-7 or SUN-SAT</td>
    /// <td align="left"></td>
    /// <td align="left">, - /// ? / L #</td>
    /// </tr>
    /// <tr>
    /// <td align="left">Year (Optional)</td>
    /// <td align="left"> </td>
    /// <td align="left">empty, <see cref="TriggerConstants.EarliestYear"/>- <see cref="TriggerConstants.YearToGiveUpSchedulingAt"/></td>
    /// <td align="left"> </td>
    /// <td align="left">, - /// /</td>
    /// </tr>
    /// </table>
    /// <para>
    /// The '*' character is used to specify all values. For example, &quot;*&quot;
    /// in the minute field means &quot;every minute&quot;.
    /// </para>
    /// <para>
    /// The '?' character is allowed for the day-of-month and day-of-week fields. It
    /// is used to specify 'no specific value'. This is useful when you need to
    /// specify something in one of the two fields, but not the other.
    /// </para>
    /// <para>
    /// The '-' character is used to specify ranges. For example &quot;10-12&quot; in
    /// the hour field means &quot;the hours 10, 11 and 12&quot;.
    /// </para>
    /// <para>
    /// The ',' character is used to specify additional values. For example
    /// &quot;MON,WED,FRI&quot; in the day-of-week field means &quot;the days Monday,
    /// Wednesday, and Friday&quot;.
    /// </para>
    /// <para>
    /// The '/' character is used to specify increments. For example &quot;0/15&quot;
    /// in the seconds field means &quot;the seconds 0, 15, 30, and 45&quot;. And
    /// &quot;5/15&quot; in the seconds field means &quot;the seconds 5, 20, 35, and
    /// 50&quot;.  Specifying '*' before the  '/' is equivalent to specifying 0 is
    /// the value to start with. Essentially, for each field in the expression, there
    /// is a set of numbers that can be turned on or off. For seconds and minutes,
    /// the numbers range from 0 to 59. For hours 0 to 23, for days of the month 1 to
    /// 31, and for months 1 to 12. The &quot;/&quot; character simply helps you turn
    /// on every &quot;nth&quot; value in the given set. Thus, &quot;7/6&quot; in the
    /// month field only turns on month &quot;7&quot;, it does NOT mean every 6th
    /// month, please note that subtlety.
    /// </para>
    /// <para>
    /// The 'L' character is allowed for the day-of-month and day-of-week fields.
    /// This character is short-hand for &quot;last&quot;, but it has different
    /// meaning in each of the two fields. For example, the value &quot;L&quot; in
    /// the day-of-month field means &quot;the last day of the month&quot; - day 31
    /// for January, day 28 for February on non-leap years. If used in the
    /// day-of-week field by itself, it simply means &quot;7&quot; or
    /// &quot;SAT&quot;. But if used in the day-of-week field after another value, it
    /// means &quot;the last xxx day of the month&quot; - for example &quot;6L&quot;
    /// means &quot;the last friday of the month&quot;. You can also specify an offset
    /// from the last day of the month, such as "L-3" which would mean the third-to-last
    /// day of the calendar month. <i>When using the 'L' option, it is important not to
    /// specify lists, or ranges of values, as you'll get confusing/unexpected results.</i>
    /// </para>
    /// <para>
    /// The 'W' character is allowed for the day-of-month field.  This character
    /// is used to specify the weekday (Monday-Friday) nearest the given day.  As an
    /// example, if you were to specify &quot;15W&quot; as the value for the
    /// day-of-month field, the meaning is: &quot;the nearest weekday to the 15th of
    /// the month&quot;. So if the 15th is a Saturday, the trigger will fire on
    /// Friday the 14th. If the 15th is a Sunday, the trigger will fire on Monday the
    /// 16th. If the 15th is a Tuesday, then it will fire on Tuesday the 15th.
    /// However, if you specify &quot;1W&quot; as the value for day-of-month, and the
    /// 1st is a Saturday, the trigger will fire on Monday the 3rd, as it will not
    /// 'jump' over the boundary of a month's days.  The 'W' character can only be
    /// specified when the day-of-month is a single day, not a range or list of days.
    /// </para>
    /// <para>
    /// The 'L' and 'W' characters can also be combined for the day-of-month
    /// expression to yield 'LW', which translates to &quot;last weekday of the
    /// month&quot;.
    /// </para>
    /// <para>
    /// The '#' character is allowed for the day-of-week field. This character is
    /// used to specify &quot;the nth&quot; XXX day of the month. For example, the
    /// value of &quot;6#3&quot; in the day-of-week field means the third Friday of
    /// the month (day 6 = Friday and &quot;#3&quot; = the 3rd one in the month).
    /// Other examples: &quot;2#1&quot; = the first Monday of the month and
    /// &quot;4#5&quot; = the fifth Wednesday of the month. Note that if you specify
    /// &quot;#5&quot; and there is not 5 of the given day-of-week in the month, then
    /// no firing will occur that month. If the '#' character is used, there can
    /// only be one expression in the day-of-week field (&quot;3#1,6#3&quot; is
    /// not valid, since there are two expressions).
    /// </para>
    /// <para>
    /// <!--The 'C' character is allowed for the day-of-month and day-of-week fields.
    /// This character is short-hand for "calendar". This means values are
    /// calculated against the associated calendar, if any. If no calendar is
    /// associated, then it is equivalent to having an all-inclusive calendar. A
    /// value of "5C" in the day-of-month field means "the first day included by the
    /// calendar on or after the 5th". A value of "1C" in the day-of-week field
    /// means "the first day included by the calendar on or after Sunday". -->
    /// </para>
    /// <para>
    /// The legal characters and the names of months and days of the week are not
    /// case-sensitive.
    /// </para>
    /// <para>
    /// <b>NOTES:</b>
    /// <ul>
    /// <li>Support for specifying both a day-of-week and a day-of-month value is
    /// not complete (you'll need to use the '?' character in one of these fields).
    /// </li>
    /// <li>Overflowing ranges are supported - that is, having a larger number on
    /// the left-hand side than the right. You might do 22-2 to catch 10 o'clock
    /// at night until 2 o'clock in the morning, or you might have NOV-FEB. It is
    /// very important to note that overuse of overflowing ranges creates ranges
    /// that don't make sense and no effort has been made to determine which
    /// interpretation CronExpression chooses. An example would be
    /// "0 0 14-6 ? * FRI-MON". </li>
    /// </ul>
    /// </para>
    /// </remarks>
    /// <author>Sharada Jambula</author>
    /// <author>James House</author>
    /// <author>Contributions from Mads Henderson</author>
    /// <author>Refactoring from CronTrigger to CronExpression by Aaron Craven</author>
    /// <author>Marko Lahma (.NET)</author>
    [Serializable]
    public sealed class CronExpression : ISerializable {
        private TimeZoneInfo? timeZone;

        [NonSerialized] private readonly CronField seconds = new();
        [NonSerialized] private readonly CronField minutes = new();
        [NonSerialized] private readonly CronField hours = new();
        [NonSerialized] private readonly CronField daysOfMonth = new();
        [NonSerialized] private readonly CronField months = new();
        [NonSerialized] private readonly CronField daysOfWeek = new();
        [NonSerialized] private readonly CronField years = new();

        /// <summary>
        /// Last day of week.
        /// </summary>
        [NonSerialized] private bool lastDayOfWeek;

        /// <summary>
        /// N number of weeks.
        /// </summary>
        [NonSerialized] private int everyNthWeek;

        /// <summary>
        /// Nth day of the week.
        /// </summary>
        [NonSerialized] private int nthdayOfWeek;

        /// <summary>
        /// Last day of month.
        /// </summary>
        [NonSerialized] private bool lastDayOfMonth;

        /// <summary>
        /// Nearest weekday.
        /// </summary>
        [NonSerialized] private bool nearestWeekday;

        [NonSerialized] private int lastDayOffset;

        [NonSerialized] private int lastWeekdayOffset;

        /// <summary>
        /// Calendar day of the week.
        /// </summary>
        [NonSerialized] private bool calendarDayOfWeek;

        /// <summary>
        /// Calendar day of the month.
        /// </summary>
        [NonSerialized] private bool calendarDayOfMonth;

        private static readonly Regex regex = new("^L(-\\d{1,2})?(W(-\\d{1,2})?)?$", RegexOptions.Compiled | RegexOptions.ExplicitCapture, TimeSpan.FromSeconds(5)); //e.g. LW L-0W L-4 L-12W LW-4 LW-12
        private static readonly Regex offsetRegex = new("LW-(?<offset>[0-9]+)", RegexOptions.Compiled | RegexOptions.ExplicitCapture, TimeSpan.FromSeconds(5));

        static CronExpression() {
        }

        ///<summary>
        /// Constructs a new <see cref="CronExpressionString" /> based on the specified
        /// parameter.
        /// </summary>
        /// <param name="cronExpression">String representation of the cron expression the new object should represent. 
        /// [sec] [min] [hr] [day] [month] [dw] [year]
        /// </param>
        /// <see cref="CronExpressionString" />
        /// <remarks>
        ///
        /// Cron reference
        /// ┌────────────── second (0 - 59)
        /// │ ┌───────────── minute (0 - 59)
        /// │ │ ┌───────────── hour (0 - 23)
        /// │ │ │ ┌───────────── day of month (1 - 31)
        /// │ │ │ │ ┌───────────── month (1 - 12)
        /// │ │ │ │ │ ┌───────────── day of week (0 - 6) (Sunday to Saturday; 7 is also Sunday on some systems)
        /// │ │ │ │ │ │ ┌──────────── year
        /// │ │ │ │ │ │ │
        /// │ │ │ │ │ │ │
        /// * * * * * * * command to execute
        ///
        /// Quartz uses different day of week definition. 1 - 7 (Sunday to Saturday). To avoid the issue, suggests to use SUN, MON, ... instead of numbers.
        ///
        /// </remarks>
        public CronExpression(string cronExpression) {
            if (cronExpression is null) {
                throw new ArgumentNullException(nameof(cronExpression));
            }

            this.CronExpressionString = CultureInfo.InvariantCulture.TextInfo.ToUpper(cronExpression).Trim();
            this.BuildExpression(this.CronExpressionString);
        }

        private static int GetVersion(SerializationInfo info) {
            try {
                return info.GetInt32("version");
            }
            catch {
                return 0;
            }
        }

        /// <summary>
        /// Serialization constructor.
        /// </summary>
        private CronExpression(SerializationInfo info, StreamingContext context) {
            var version = GetVersion(info);
            switch (version) {
                case 0:
                    this.CronExpressionString = (string)info.GetValue("cronExpressionString", typeof(string))!;
                    this.TimeZone = (TimeZoneInfo)info.GetValue("timeZone", typeof(TimeZoneInfo))!;
                    break;
                case 1:
                    this.CronExpressionString = (string)info.GetValue("cronExpression", typeof(string))!;
                    var timeZoneId = (string)info.GetValue("timeZoneId", typeof(string))!;
                    if (!string.IsNullOrEmpty(timeZoneId)) {
                        this.timeZone = TimeZoneUtil.FindTimeZoneById(timeZoneId);
                    }

                    break;
                default:
                    throw new NotSupportedException($"Unknown serialization version {version}");
                    break;
            }

            this.BuildExpression(this.CronExpressionString);
        }

        [System.Security.SecurityCritical]
        public void GetObjectData(SerializationInfo info, StreamingContext context) {
            info.AddValue("version", 1);
            info.AddValue("cronExpression", this.CronExpressionString);
            info.AddValue("timeZoneId", this.TimeZone.Id);
        }

        /// <summary>
        /// Indicates whether the given date satisfies the cron expression.
        /// </summary>
        /// <remarks>
        /// Note that  milliseconds are ignored, so two Dates falling on different milliseconds
        /// of the same second will always have the same result here.
        /// </remarks>
        /// <param name="dateUtc">The date to evaluate.</param>
        /// <returns>a boolean indicating whether the given date satisfies the cron expression</returns>
        public bool IsSatisfiedBy(DateTimeOffset dateUtc) {
            var withoutMilliseconds = new DateTimeOffset(dateUtc.Year, dateUtc.Month, dateUtc.Day, dateUtc.Hour, dateUtc.Minute, dateUtc.Second, dateUtc.Offset);
            var test = withoutMilliseconds.AddSeconds(-1);
            var timeAfter = this.GetTimeAfter(test);

            return timeAfter.HasValue
                   && timeAfter.Value.Equals(withoutMilliseconds);
        }

        /// <summary>
        /// Returns the next date/time <i>after</i> the given date/time which
        /// satisfies the cron expression.
        /// </summary>
        /// <param name="date">the date/time at which to begin the search for the next valid date/time</param>
        /// <returns>the next valid date/time</returns>
        public DateTimeOffset? GetNextValidTimeAfter(DateTimeOffset date) {
            return this.GetTimeAfter(date);
        }

        /// <summary>
        /// Returns the next date/time <i>after</i> the given date/time which does
        /// <i>not</i> satisfy the expression.
        /// </summary>
        /// <param name="date">the date/time at which to begin the search for the next invalid date/time</param>
        /// <returns>the next valid date/time</returns>
        public DateTimeOffset? GetNextInvalidTimeAfter(DateTimeOffset date) {
            long difference = 1000;

            // move back to the nearest second so differences will be accurate
            var lastDate = new DateTimeOffset(date.Year, date.Month, date.Day, date.Hour, date.Minute, date.Second, date.Offset).AddSeconds(-1);

            //TODO: IMPROVE THIS! The following is a BAD solution to this problem. Performance will be very bad here, depending on the cron expression. It is, however A solution.

            // Keep getting the next included time until it's farther than one second
            // apart. At that point, lastDate is the last valid fire time. We return
            // the second immediately following it.
            while (difference == 1000) {
                var newDate = this.GetTimeAfter(lastDate);

                if (newDate is null) {
                    break;
                }

                difference = (long)(newDate.Value - lastDate).TotalMilliseconds;

                if (difference == 1000) {
                    lastDate = newDate.Value;
                }
            }

            return lastDate.AddSeconds(1);
        }

        /// <summary>
        /// Sets or gets the time zone for which the <see cref="CronExpression" /> of this
        /// <see cref="ICronTrigger" /> will be resolved.
        /// </summary>
        public TimeZoneInfo TimeZone {
            set => this.timeZone = value;
            get => this.timeZone ??= TimeZoneInfo.Local;
        }

        /// <summary>
        /// Returns the string representation of the <see cref="CronExpression" />
        /// </summary>
        /// <returns>The string representation of the <see cref="CronExpression" /></returns>
        public override string ToString() {
            return this.CronExpressionString;
        }

        /// <summary>
        /// Indicates whether the specified cron expression can be parsed into a
        /// valid cron expression
        /// </summary>
        /// <param name="cronExpression">the expression to evaluate</param>
        /// <returns>a boolean indicating whether the given expression is a valid cron
        ///         expression</returns>
        public static bool IsValidExpression(string cronExpression) {
            try {
                var _ = new CronExpression(cronExpression);
            }
            catch (FormatException) {
                return false;
            }

            return true;
        }

        public static void ValidateExpression(string cronExpression) {
            var _ = new CronExpression(cronExpression);
        }

        ////////////////////////////////////////////////////////////////////////////
        //
        // Expression Parsing Functions
        //
        ////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Builds the expression.
        /// </summary>
        /// <param name="expression">The expression.</param>
        private void BuildExpression(string expression) {
            try {
                this.ClearExpressionFields();

                var exprOn = CronExpressionConstants.Second;

                foreach (var (expr, _) in expression.SpanSplit(' ', '\t')) {
                    if (exprOn > CronExpressionConstants.Year) {
                        break;
                    }

                    // throw an exception if L is used with other days of the month
                    if (exprOn == CronExpressionConstants.DayOfMonth) {
                        if (expr.IndexOf('L') != -1 && expr.Length > 1 && expr.IndexOf(',') >= 0 && expr.Slice(expr.IndexOf('L') + 1).IndexOf('L') != -1) {
                            throw new FormatException("Support for specifying 'L' with other days of the month is limited to one instance of L");
                        }
                    }

                    // throw an exception if L is used with other days of the week
                    if (exprOn == CronExpressionConstants.DayOfWeek && expr.IndexOf('L') != -1 && expr.Length > 1 && expr.IndexOf(',') >= 0) {
                        throw new FormatException("Support for specifying 'L' with other days of the week is not implemented");
                    }

                    if (exprOn == CronExpressionConstants.DayOfWeek && expr.IndexOf('#') != -1 && expr.Slice(expr.IndexOf('#') + 1 + 1).IndexOf('#') != -1) {
                        throw new FormatException("Support for specifying multiple \"nth\" days is not implemented.");
                    }

                    if (expr.IndexOf(',') != -1) {
                        foreach (var v in expr.SpanSplit(',')) {
                            this.StoreExpressionValues(0, v, exprOn);
                        }
                    }
                    else {
                        // simple field
                        this.StoreExpressionValues(0, expr, exprOn);
                    }

                    exprOn++;
                }

                if (exprOn <= CronExpressionConstants.DayOfWeek) {
                    throw new FormatException("Unexpected end of expression.");
                }

                if (exprOn <= CronExpressionConstants.Year) {
                    this.StoreExpressionValues(0, "*".AsSpan(), CronExpressionConstants.Year);
                }
            }
            catch (FormatException) {
                throw;
            }
            catch (Exception e) {
                throw new FormatException($"Illegal cron expression format ({e.Message})", e);
            }
        }

        private void ClearExpressionFields() {
            this.seconds.Clear();
            this.minutes.Clear();
            this.hours.Clear();
            this.daysOfMonth.Clear();
            this.months.Clear();
            this.daysOfWeek.Clear();
            this.years.Clear();
        }

        private void StoreExpressionQuestionMark(int type, ReadOnlySpan<char> s, int i) {
            i++;
            if (i + 1 <= s.Length && !char.IsWhiteSpace(s[i])) {
                throw new FormatException("Illegal character after '?': " + s[i]);
            }

            if (type != CronExpressionConstants.DayOfWeek && type != CronExpressionConstants.DayOfMonth) {
                throw new FormatException("'?' can only be specified for Day-of-Month or Day-of-Week.");
            }

            if (type == CronExpressionConstants.DayOfWeek && !this.lastDayOfMonth) {
                var val = this.daysOfMonth.LastOrDefault();
                if (val == CronExpressionConstants.NoSpec) {
                    throw new FormatException("'?' can only be specified for Day-of-Month -OR- Day-of-Week.");
                }
            }

            this.AddToSet(CronExpressionConstants.NoSpec, -1, 0, type);
        }

        private void StoreExpressionStarOrSlash(int type, ReadOnlySpan<char> s, int i) {
            var c = s[i];
            var incr = 0;
            var startsWithAsterisk = c == '*';
            if (startsWithAsterisk && i + 1 >= s.Length) {
                this.AddToSet(CronExpressionConstants.AllSpec, -1, incr, type);
                return;
            }

            if (c == '/' && (i + 1 >= s.Length || char.IsWhiteSpace(s[i + 1]))) {
                throw new FormatException("'/' must be followed by an integer.");
            }

            if (startsWithAsterisk) {
                i++;
            }

            c = s[i];
            if (c == '/') {
                // is an increment specified?
                i++;
                if (i >= s.Length) {
                    throw new FormatException("Unexpected end of string.");
                }

                incr = CronExpression.GetNumericValue(s, i);
                CheckIncrementRange(incr, type);
            }
            else {
                if (startsWithAsterisk) {
                    throw new FormatException("Illegal characters after asterisk: " + s.ToString());
                }

                incr = 1;
            }

            this.AddToSet(CronExpressionConstants.AllSpec, -1, incr, type);
        }

        private void StoreExpressionL(int type, ReadOnlySpan<char> s, int i) {
            i++;
            switch (type) {
                case CronExpressionConstants.DayOfMonth: {
                        this.lastDayOfMonth = true;
                        if (s.Length > i) {
                            var c = s[i];
                            if (c == '-') {
                                (this.lastDayOffset, i) = GetValue(0, s, i + 1);
                                if (this.lastDayOffset > 30) {
                                    throw new FormatException("Offset from last day must be <= 30");
                                }
                            }

                            if (s.Length > i) {
                                c = s[i];
                                if (c == 'W') {
                                    this.nearestWeekday = true;
                                }

                                var match = offsetRegex.Match(s.ToString());
                                if (match.Success) {
                                    var offSetGroup = match.Groups["offset"];
                                    if (offSetGroup.Success) {
                                        this.lastWeekdayOffset = int.Parse(offSetGroup.Value);
                                    }
                                }
                            }
                        }

                        break;
                    }

                case CronExpressionConstants.DayOfWeek:
                    this.AddToSet(7, 7, 0, type);
                    break;
                default:
                    throw new FormatException($"'L' option is not valid here. (pos={i})");
                    break;
            }
        }

        private void StoreExpressionNumeric(int type, ReadOnlySpan<char> s, int i) {
            // test fast case
            if (int.TryParse(s, out var temp)) {
                this.AddToSet(temp, -1, -1, type);
                return;
            }

            var c = s[i];
            var val = ToInt32(c);
            i++;
            if (i >= s.Length) {
                this.AddToSet(val, -1, -1, type);
            }
            else {
                c = s[i];
                if (char.IsDigit(c)) {
                    (val, i) = GetValue(val, s, i);
                }

                this.CheckNext(i, s, val, type);
            }
        }

        private void StoreExpressionGeneralValue(int type, ReadOnlySpan<char> s, int i) {
            var incr = 0;
            var sub = s.Slice(i, 3);
            int sval;
            var eval = -1;
            if (type == CronExpressionConstants.Month) {
                sval = CronExpression.GetMonthNumber(sub) + 1;
                if (sval <= 0) {
                    throw new FormatException($"Invalid Month value: '{sub.ToString()}'");
                }

                if (s.Length > i + 3) {
                    if (s[i + 3] == '-') {
                        i += 4;
                        sub = s.Slice(i, 3);
                        eval = GetMonthNumber(sub) + 1;
                        if (eval <= 0) {
                            throw new FormatException($"Invalid Month value: '{sub.ToString()}'");
                        }
                    }
                }
            }
            else if (type == CronExpressionConstants.DayOfWeek) {
                sval = GetDayOfWeekNumber(sub);
                if (sval < 0) {
                    throw new FormatException($"Invalid Day-of-Week value: '{sub.ToString()}'");
                }

                if (s.Length > i + 3) {
                    var c = s[i + 3];
                    switch (c) {
                        case '-':
                            i += 4;
                            sub = s.Slice(i, 3);
                            eval = GetDayOfWeekNumber(sub);
                            if (eval < 0) {
                                throw new FormatException($"Invalid Day-of-Week value: '{sub.ToString()}'");
                            }

                            break;
                        case '#':
                            try {
                                i += 4;
                                this.nthdayOfWeek = ToInt32(s.Slice(i));
                                if (this.nthdayOfWeek is < 1 or > 5) {
                                    throw new FormatException("nthdayOfWeek is < 1 or > 5");
                                }
                            }
                            catch (Exception) {
                                throw new FormatException("A numeric value between 1 and 5 must follow the '#' option");
                            }

                            break;
                        case '/':
                            try {
                                i += 4;
                                this.everyNthWeek = ToInt32(s.Slice(i));
                                if (this.everyNthWeek is < 1 or > 5) {
                                    throw new FormatException("everyNthWeek is < 1 or > 5");
                                }
                            }
                            catch (Exception) {
                                throw new FormatException("A numeric value between 1 and 5 must follow the '/' option");
                            }

                            break;
                        case 'L':
                            this.lastDayOfWeek = true;
                            break;
                        default:
                            throw new FormatException($"Illegal characters for this position: '{sub.ToString()}'");
                            break;
                    }
                }
            }
            else {
                throw new FormatException($"Illegal characters for this position: '{sub.ToString()}'");
                return;
            }

            if (eval != -1) {
                incr = 1;
            }

            this.AddToSet(sval, eval, incr, type);
        }

        private void StoreExpressionValues(int pos, ReadOnlySpan<char> s, int type) {
            var i = pos;
            if (i < s.Length && char.IsWhiteSpace(s[i])) {
                i = SkipWhiteSpace(pos, s);
            }

            if (i >= s.Length) {
                return;
            }

            switch (s[i]) {
                case >= 'A' and <= 'Z' when !s.SequenceEqual("L".AsSpan()) && !regex.IsMatch(s.ToString()):
                    this.StoreExpressionGeneralValue(type, s, i);
                    break;

                case '?':
                    this.StoreExpressionQuestionMark(type, s, i);
                    break;

                case '*':
                case '/':
                    this.StoreExpressionStarOrSlash(type, s, i);
                    break;

                case 'L':
                    this.StoreExpressionL(type, s, i);
                    break;

                case >= '0' and <= '9':
                    this.StoreExpressionNumeric(type, s, i);
                    break;
                default:
                    throw new FormatException($"Unexpected character: {s[i]}");
                    break;
            }
        }

        // ReSharper disable once UnusedParameter.Local
        private static void CheckIncrementRange(int incr, int type) {
            switch (type) {
                case CronExpressionConstants.Second or CronExpressionConstants.Minute when incr > 59:
                    throw new FormatException($"Increment > 59 : {incr}");
                    break;
                case CronExpressionConstants.Hour when incr > 23:
                    throw new FormatException($"Increment > 23 : {incr}");
                    break;
                case CronExpressionConstants.DayOfMonth when incr > 31:
                    throw new FormatException($"Increment > 31 : {incr}");
                    break;
                case CronExpressionConstants.DayOfWeek when incr > 7:
                    throw new FormatException($"Increment > 7 : {incr}");
                    break;
                case CronExpressionConstants.Month when incr > 12:
                    throw new FormatException($"Increment > 12 : {incr}");
                    break;
            }
        }

        private void CheckNext(int pos, ReadOnlySpan<char> s, int val, int type) {
            if (pos >= s.Length) {
                this.AddToSet(val, -1, -1, type);
                return;
            }

            switch (s[pos]) {
                case 'L':
                    this.HandleLOption(val, type, pos);
                    return;

                case 'W':
                    this.HandleWOption(val, type, pos);
                    return;

                case '#':
                    this.HandleHashOption(s, val, type, pos);
                    return;

                case 'C':
                    this.HandleCOption(val, type, pos);
                    return;

                case '-':
                    this.HandleDashOption(s, val, type, pos);
                    return;

                case '/':
                    this.HandleSlashOption(s, val, type, pos, -1);
                    return;

                default:
                    this.AddToSet(val, -1, 0, type);
                    return;
            }
        }

        private void HandleSlashOption(ReadOnlySpan<char> s, int val, int type, int i, int end) {
            if (i + 1 >= s.Length || char.IsWhiteSpace(s[i + 1])) {
                throw new FormatException("\'/\' must be followed by an integer.");
            }

            i++;
            var c = s[i];
            var v2 = ToInt32(c);
            i++;
            if (i >= s.Length) {
                CheckIncrementRange(v2, type);
                this.AddToSet(val, end, v2, type);
                return;
            }

            c = s[i];
            if (char.IsDigit(c)) {
                var (v3, _) = GetValue(v2, s, i);
                CheckIncrementRange(v3, type);
                this.AddToSet(val, end, v3, type);
                return;
            }

            throw new FormatException($"Unexpected character '{c}' after '/'");
        }

        private void HandleDashOption(ReadOnlySpan<char> s, int val, int type, int i) {
            i++;
            var c = s[i];
            var v = ToInt32(c);
            var end = v;
            i++;
            if (i >= s.Length) {
                this.AddToSet(val, end, 1, type);
                return;
            }

            c = s[i];
            if (char.IsDigit(c)) {
                (end, i) = GetValue(v, s, i);
            }

            if (i < s.Length && s[i] == '/') {
                i++;
                c = s[i];
                var v2 = ToInt32(c);
                i++;
                if (i >= s.Length) {
                    this.AddToSet(val, end, v2, type);
                    return;
                }

                c = s[i];
                if (char.IsDigit(c)) {
                    var (v3, _) = GetValue(v2, s, i);
                    this.AddToSet(val, end, v3, type);
                    return;
                }

                this.AddToSet(val, end, v2, type);
                return;
            }

            this.AddToSet(val, end, 1, type);
        }

        private void HandleCOption(int val, int type, int i) {
            switch (type) {
                case CronExpressionConstants.DayOfWeek:
                    this.calendarDayOfWeek = true;
                    break;
                case CronExpressionConstants.DayOfMonth:
                    this.calendarDayOfMonth = true;
                    break;
                default:
                    throw new FormatException($"'C' option is not valid here. (pos={i})");
                    break;
            }

            var data = this.GetSet(type);
            data.Add(val);
        }

        private void HandleHashOption(ReadOnlySpan<char> s, int val, int type, int i) {
            var pos = i;
            if (type != CronExpressionConstants.DayOfWeek) {
                throw new FormatException($"'#' option is not valid here. (pos={i})");
            }

            i++;
            try {
                this.nthdayOfWeek = ToInt32(s.Slice(i));
                if (this.nthdayOfWeek is < 1 or > 5) {
                    throw new FormatException("nthdayOfWeek is < 1 or > 5");
                }

                // check first char is numeric and is a valid Day of week (1-7)
                if (int.TryParse(s.Slice(0, pos), out val)) {
                    if (val is < 1 or > 7) {
                        throw new FormatException("Day-of-Week values must be between 1 and 7");
                    }
                }
            }
            catch (Exception) {
                throw new FormatException("A numeric value between 1 and 5 must follow the '#' option");
            }

            var data = this.GetSet(type);
            data.Add(val);
        }

        private void HandleWOption(int val, int type, int i) {
            if (type == CronExpressionConstants.DayOfMonth) {
                this.nearestWeekday = true;
            }
            else {
                throw new FormatException($"'W' option is not valid here. (pos={i})");
            }

            if (val > 31) {
                throw new FormatException("The 'W' option does not make sense with values larger than 31 (max number of days in a month)");
            }

            var data = this.GetSet(type);
            data.Add(val);
        }

        private void HandleLOption(int val, int type, int pos) {
            if (type == CronExpressionConstants.DayOfWeek) {
                if (val is < 1 or > 7) {
                    throw new FormatException("Day-of-Week values must be between 1 and 7");
                }

                this.lastDayOfWeek = true;
            }
            else {
                throw new FormatException($"'L' option is not valid here. (pos={pos})");
            }

            var data = this.GetSet(type);
            data.Add(val);
        }

        /// <summary>
        /// Gets the cron expression string.
        /// [sec] [min] [hr] [day] [month] [dw] [year]
        /// </summary>
        /// <example>
        /// 
        /// Cron reference
        /// ┌────────────── second (0 - 59)
        /// │ ┌───────────── minute (0 - 59)
        /// │ │ ┌───────────── hour (0 - 23)
        /// │ │ │ ┌───────────── day of month (1 - 31)
        /// │ │ │ │ ┌───────────── month (1 - 12)
        /// │ │ │ │ │ ┌───────────── day of week (0 - 6) (Sunday to Saturday; 7 is also Sunday on some systems)
        /// │ │ │ │ │ │ ┌──────────── year
        /// │ │ │ │ │ │ │
        /// │ │ │ │ │ │ │
        /// * * * * * * * command to execute
        ///
        ///
        /// </example>
        /// <value>The cron expression string.</value>
        public string CronExpressionString { get; }

        /// <summary>
        /// Gets the expression summary.
        /// </summary>
        /// <returns></returns>
        public string GetExpressionSummary() {
            return new CronExpressionSummary(
                this.seconds,
                this.minutes,
                this.hours,
                this.daysOfMonth,
                this.months,
                this.daysOfWeek,
                this.lastDayOfWeek,
                this.nearestWeekday,
                this.nthdayOfWeek,
                this.lastDayOfMonth,
                this.calendarDayOfWeek,
                this.calendarDayOfMonth,
                this.years
            ).ToString();
        }

        private static int SkipWhiteSpace(int position, ReadOnlySpan<char> str) {
            for (; position < str.Length && char.IsWhiteSpace(str[position]); position++) {
            }

            return position;
        }

        private static int FindNextWhiteSpace(int position, ReadOnlySpan<char> str) {
            for (; position < str.Length && !char.IsWhiteSpace(str[position]); position++) {
            }

            return position;
        }

        private static (int min, int max, string errorMessage) GetValidationParameters(int type) {
            return type switch {
                CronExpressionConstants.Second or CronExpressionConstants.Minute
                    => (0, 59, "Minute and Second values must be between 0 and 59"),
                CronExpressionConstants.Hour
                    => (0, 23, "Hour values must be between 0 and 23"),
                CronExpressionConstants.DayOfMonth
                    => (1, 31, "Day of month values must be between 1 and 31"),
                CronExpressionConstants.Month
                    => (1, 12, "Month values must be between 1 and 12"),
                CronExpressionConstants.DayOfWeek
                    => (1, 7, "Day-of-Week values must be between 1 and 7"),
                CronExpressionConstants.Year
                    => (TriggerConstants.EarliestYear, TriggerConstants.YearToGiveUpSchedulingAt, $"Year values must be between {TriggerConstants.EarliestYear} and {TriggerConstants.YearToGiveUpSchedulingAt}"),
                _ => throw new ArgumentOutOfRangeException(nameof(type), "Invalid cron expression type")
            };
        }

        private static bool IsSpecialValue(int val, int type) {
            return val == CronExpressionConstants.AllSpec ||
                   (type is CronExpressionConstants.DayOfMonth or CronExpressionConstants.DayOfWeek &&
                    val == CronExpressionConstants.NoSpec);
        }

        private static void ValidateSetValues(int val, int end, int type) {
            var (min, max, errorMessage) = GetValidationParameters(type);

            if ((val < min || val > max || end > max) &&
                !IsSpecialValue(val, type)) {
                throw new FormatException(errorMessage);
            }
        }

        private static (int startAt, int stopAt) GetRangeForType(int type, int val, int end) {
            return type switch {
                CronExpressionConstants.Second or CronExpressionConstants.Minute => (GetStartAt(val, 0), GetStopAt(end, 59)),
                CronExpressionConstants.Hour => (GetStartAt(val, 0), GetStopAt(end, 23)),
                CronExpressionConstants.DayOfMonth => (GetStartAt(val, 1), GetStopAt(end, 31)),
                CronExpressionConstants.Month => (GetStartAt(val, 1), GetStopAt(end, 12)),
                CronExpressionConstants.DayOfWeek => (GetStartAt(val, 1), GetStopAt(end, 7)),
                CronExpressionConstants.Year => (GetStartAt(val, TriggerConstants.EarliestYear), GetStopAt(end, TriggerConstants.YearToGiveUpSchedulingAt)),
                _ => throw new ArgumentException("Unexpected type encountered")
            };
        }

        /// <summary>
        /// Gets the max value for the cron expression type.
        /// </summary>
        /// <param name="type"> The type of the cron expression</param>
        /// <param name="startAt"> The start value</param>
        /// <param name="stopAt"> The stop value</param>
        /// <returns>Returns -1 if stopAt is less than startAt otherwise returns the max value for the type</returns>
        private static int GetMaxValueForType(int type, int startAt, int stopAt) {
            if (stopAt >= startAt) return -1;

            return type switch {
                CronExpressionConstants.Second or CronExpressionConstants.Minute => 60,
                CronExpressionConstants.Hour => 24,
                CronExpressionConstants.Month => 12,
                CronExpressionConstants.DayOfWeek => 7,
                CronExpressionConstants.DayOfMonth => 31,
                CronExpressionConstants.Year => throw new ArgumentException("Start year must be less than stop year"),
                _ => throw new ArgumentException("Unexpected type encountered")
            };
        }

        private static int GetStartAt(int val, int defaultValue) =>
            val is -1 or CronExpressionConstants.AllSpec ? defaultValue : val;

        private static int GetStopAt(int end, int defaultValue) =>
            end == -1 ? defaultValue : end;

        private void AddToSet(int val, int end, int incr, int type) {
            ValidateSetValues(val, end, type);

            var data = this.GetSet(type);

            if (incr is 0 or -1 && val != CronExpressionConstants.AllSpec) {
                data.Add(val != -1 ? val : CronExpressionConstants.NoSpec);
                return;
            }

            if (val == CronExpressionConstants.AllSpec && incr <= 0) {
                data.Add(CronExpressionConstants.AllSpec);
                // skip adding other data, we check this wildcard in TryGetMinValueStartingFrom
                return;
            }

            var (startAt, stopAt) = GetRangeForType(type, val, end);

            // if the end of the range is before the start, then we need to overflow into
            // the next day, month etc. This is done by adding the maximum amount for that
            // type, and using modulus max to determine the value being added.
            var max = GetMaxValueForType(type, startAt, stopAt);
            if (max != -1)
                stopAt += max;

            for (var i = startAt; i <= stopAt; i += incr) {
                if (max == -1) {
                    // ie: there's no max to overflow over
                    data.Add(i);
                }
                else {
                    // take the modulus to get the real value
                    var i2 = i % max;

                    // 1-indexed ranges should not include 0, and should include their max
                    if (i2 == 0 && (type == CronExpressionConstants.Month
                                    || type == CronExpressionConstants.DayOfWeek
                                    || type == CronExpressionConstants.DayOfMonth)) {
                        i2 = max;
                    }

                    data.Add(i2);
                }
            }
        }

        /// <summary>
        /// Gets the set of given type.
        /// </summary>
        internal CronField GetSet(int type) {
            var field = type switch {
                CronExpressionConstants.Second => this.seconds,
                CronExpressionConstants.Minute => this.minutes,
                CronExpressionConstants.Hour => this.hours,
                CronExpressionConstants.DayOfMonth => this.daysOfMonth,
                CronExpressionConstants.Month => this.months,
                CronExpressionConstants.DayOfWeek => this.daysOfWeek,
                CronExpressionConstants.Year => this.years,
                _ => default
            };

            if (field is null) {
                throw new ArgumentOutOfRangeException(nameof(type));
            }

            return field;
        }

        private static ValueAndPosition GetValue(int v, ReadOnlySpan<char> s, int i) {
            var c = s[i];

            var builder = new StringBuilder(s.Length);
            builder.Append(v);

            while (char.IsDigit(c)) {
                builder.Append(c);
                i++;
                if (i >= s.Length) {
                    break;
                }

                c = s[i];
            }

            var value = Convert.ToInt32(builder.ToString(), CultureInfo.InvariantCulture);
            var pos = i < s.Length ? i : i + 1;
            return new ValueAndPosition(value, pos);
        }

        /// <summary>
        /// Gets the numeric value from string.
        /// </summary>
        private static int GetNumericValue(ReadOnlySpan<char> s, int i) {
            var endOfVal = CronExpression.FindNextWhiteSpace(i, s);
            return ToInt32(s.Slice(i, endOfVal - i));
        }

        /// <summary>
        /// Gets the month number.
        /// </summary>
        /// <param name="s">The string to map with.</param>
        /// <returns></returns>
        private static int GetMonthNumber(ReadOnlySpan<char> s) {
            return s switch {
                "JAN" => 0,
                "FEB" => 1,
                "MAR" => 2,
                "APR" => 3,
                "MAY" => 4,
                "JUN" => 5,
                "JUL" => 6,
                "AUG" => 7,
                "SEP" => 8,
                "OCT" => 9,
                "NOV" => 10,
                "DEC" => 11,
                _ => -1
            };
        }

        private static int GetDayOfWeekNumber(ReadOnlySpan<char> s) {
            return s switch {
                "SUN" => 1,
                "MON" => 2,
                "TUE" => 3,
                "WED" => 4,
                "THU" => 5,
                "FRI" => 6,
                "SAT" => 7,
                _ => -1
            };
        }

        /// <summary>
        /// Progress next fire time seconds
        /// </summary>
        private NextFireTimeCursor ProgressNextFireTimeSecond(DateTimeOffset d) {
            var sec = d.Second;
            if (this.seconds.TryGetMinValueStartingFrom(sec, out var min)) {
                sec = min;
            }
            else {
                sec = this.seconds.Min;
                d = d.AddMinutes(1);
            }

            return new NextFireTimeCursor(false, new DateTimeOffset(d.Year, d.Month, d.Day, d.Hour, d.Minute, sec, d.Millisecond, d.Offset));
        }

        /// <summary>
        /// Progress next Fire time Minutes
        /// </summary>
        /// <param name="d">NextFireTimeCheck</param>
        private NextFireTimeCursor ProgressNextFireTimeMinute(DateTimeOffset d) {
            var minute = d.Minute;
            var hr = d.Hour;
            var t = -1;

            if (this.minutes.TryGetMinValueStartingFrom(minute, out var min)) {
                t = minute;
                minute = min;
            }
            else {
                minute = this.minutes.Min;
                hr++;
            }

            if (minute != t) {
                d = new DateTimeOffset(d.Year, d.Month, d.Day, d.Hour, minute, 0, d.Millisecond, d.Offset);
                d = SetCalendarHour(d, hr);
                return new NextFireTimeCursor(true, d);
            }

            return new NextFireTimeCursor(false, new DateTimeOffset(d.Year, d.Month, d.Day, d.Hour, minute, d.Second, d.Millisecond, d.Offset));
        }

        /// <summary>
        /// Progress next fire time Hour
        /// </summary>
        /// <param name="d">NextFireTimeCheck</param>
        private NextFireTimeCursor ProgressNextFireTimeHour(DateTimeOffset d) {
            int hour;
            var day = d.Day;
            var t = -1;

            if (this.hours.TryGetMinValueStartingFrom(d.Hour, out var min)) {
                t = d.Hour;
                hour = min;
            }
            else {
                hour = this.hours.Min;
                day++;
            }

            if (hour != t) {
                var daysInMonth = DateTime.DaysInMonth(d.Year, d.Month);
                if (day > daysInMonth) {
                    d = new DateTimeOffset(d.Year, d.Month, daysInMonth, d.Hour, 0, 0, d.Millisecond, d.Offset).AddDays(day - daysInMonth);
                }
                else {
                    d = new DateTimeOffset(d.Year, d.Month, day, d.Hour, 0, 0, d.Millisecond, d.Offset);
                }

                d = SetCalendarHour(d, hour);
                return new NextFireTimeCursor(true, d);
            }

            return new NextFireTimeCursor(false, new DateTimeOffset(d.Year, d.Month, d.Day, hour, d.Minute, d.Second, d.Millisecond, d.Offset));
        }

        private (SortedSet<int> daysOfMonthSet, bool dayHasNegativeOffset) CalculateDaysOfMonth(DateTimeOffset currentDate) {
            var daysOfMonthSet = new SortedSet<int>(this.daysOfMonth);
            var dayHasNegativeOffset = false;

            if (this.lastDayOfMonth) {
                var lastDayOfMonthValue = GetLastDayOfMonth(currentDate.Month, currentDate.Year);
                var lastDayOfMonthWithOffset = lastDayOfMonthValue - this.lastDayOffset;

                if (this.nearestWeekday) {
                    var calculatedLastDay = this.CalculateNearestWeekdayForLastDay(currentDate, lastDayOfMonthWithOffset);
                    daysOfMonthSet.Add(calculatedLastDay);
                }
                else {
                    daysOfMonthSet.Add(lastDayOfMonthWithOffset);
                }
            }
            else if (this.nearestWeekday) {
                (daysOfMonthSet, dayHasNegativeOffset) = CalculateNearestWeekdayForDaysOfMonth(currentDate, daysOfMonthSet);
            }

            return (daysOfMonthSet, dayHasNegativeOffset);
        }

        /// <summary>
        /// Calculates the nearest weekday for the last day of the month.
        /// </summary>
        /// <param name="currentDate">The current date.</param>
        /// <param name="lastDayOfMonthWithOffset">The last day of the month with the offset applied.</param>
        /// <returns>The calculated last day of the month, adjusted to the nearest weekday.</returns>
        private int CalculateNearestWeekdayForLastDay(DateTimeOffset currentDate, int lastDayOfMonthWithOffset) {
            var checkDay = new DateTimeOffset(currentDate.Year, currentDate.Month, lastDayOfMonthWithOffset, currentDate.Hour, currentDate.Minute, currentDate.Second, currentDate.Millisecond, currentDate.Offset);
            var calculatedDay = lastDayOfMonthWithOffset;

            switch (checkDay.DayOfWeek) {
                case DayOfWeek.Saturday:
                    calculatedDay -= 1;
                    break;
                case DayOfWeek.Sunday:
                    calculatedDay -= 2;
                    break;
            }

            var calculatedLastDayWithOffset = calculatedDay - this.lastWeekdayOffset;

            // If the day has crossed to the prior month, reset to 1st.
            if (calculatedLastDayWithOffset <= 0) {
                calculatedLastDayWithOffset = 1;
            }

            return calculatedLastDayWithOffset;
        }

        /// <summary>
        /// Calculates the nearest weekday for the specified days of the month.
        /// </summary>
        /// <param name="currentDate">The current date.</param>
        /// <param name="daysOfMonthSet">The set of days of the month.</param>
        /// <returns>A tuple containing the updated set of days of the month and a flag indicating if any day has a negative offset.</returns>
        private static (SortedSet<int> daysOfMonthSet, bool dayHasNegativeOffset) CalculateNearestWeekdayForDaysOfMonth(DateTimeOffset currentDate, SortedSet<int> daysOfMonthSet) {
            var endDayOfMonth = GetLastDayOfMonth(currentDate.Month, currentDate.Year);
            var minDay = (daysOfMonthSet.Min > endDayOfMonth) ? endDayOfMonth : daysOfMonthSet.Min;

            var firstDayOfMonth = new DateTimeOffset(currentDate.Year, currentDate.Month, minDay, 0, 0, 0, currentDate.Offset);
            var dayOfWeek = firstDayOfMonth.DayOfWeek;

            // Evict the original date if it is not a weekday
            if (dayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) {
                daysOfMonthSet.Remove(minDay);
            }

            var (adjustedDay, dayHasNegativeOffset) = AdjustDayToNearestWeekday(minDay, dayOfWeek, endDayOfMonth);
            daysOfMonthSet.Add(adjustedDay);

            return (daysOfMonthSet, dayHasNegativeOffset);
        }

        /// <summary>
        /// Adjusts the day to the nearest weekday based on the specified day of the week and the last day of the month.
        /// </summary>
        /// <param name="day">The day to adjust.</param>
        /// <param name="dayOfWeek">The day of the week.</param>
        /// <param name="endDayOfMonth">The end day of the month.</param>
        /// <returns>The adjusted day and a flag to indicate adjust day has a negative offset</returns>
        private static (int day, bool dayHasNegativeOffset) AdjustDayToNearestWeekday(int day, DayOfWeek dayOfWeek, int endDayOfMonth) {
            var dayHasNegativeOffset = false;

            // evict original date since it has a weekDayModifier
            switch (dayOfWeek) {
                case DayOfWeek.Saturday when day == 1:
                    day += 2;
                    break;
                case DayOfWeek.Saturday:
                    day -= 1;
                    dayHasNegativeOffset = true;
                    break;
                case DayOfWeek.Sunday when day == endDayOfMonth:
                    day -= 2;
                    dayHasNegativeOffset = true;
                    break;
                case DayOfWeek.Sunday:
                    day += 1;
                    break;
            }

            return (day, dayHasNegativeOffset);
        }

        private NextFireTimeCursor ProgressNextFireTimeDayOfMonth(DateTimeOffset d) {
            var day = d.Day;
            var month = d.Month;
            var t = -1;
            var tmon = month;

            // get day by day of month rule
            var (daysOfMonthCalculated, setIncludesDayBeforeStartDay) = this.CalculateDaysOfMonth(d);
            if (daysOfMonthCalculated.TryGetMinValueStartingFrom(d, setIncludesDayBeforeStartDay, out var min)) {
                t = day;
                day = min;

                // make sure we don't over-run a short month, such as february
                var lastDay = GetLastDayOfMonth(month, d.Year);
                if (day > lastDay) {
                    day = daysOfMonthCalculated.Min;
                    month++;
                }
            }
            else {
                if (this.lastDayOfMonth) {
                    day = daysOfMonthCalculated.Min; //for lastDayOfMonth use calculated fields
                }
                else {
                    day = this.daysOfMonth.Min; //if not, then initial set of days uncalculated (to avoid issue with stale weekday in wrong month value)
                }

                month++;
            }

            if (day != t || month != tmon) {
                if (month > 12) {
                    d = new DateTimeOffset(d.Year, 12, day, 0, 0, 0, d.Offset).AddMonths(month - 12);
                }
                else {
                    // This is to avoid a bug when moving from a month
                    // with 30 or 31 days to a month with less. Causes an invalid datetime to be instantiated.
                    // ex. 0 29 0 30 1 ? 2009 with the clock set to 1/30/2009
                    var daysInMonth = DateTime.DaysInMonth(d.Year, month);
                    if (day <= daysInMonth) {
                        d = new DateTimeOffset(d.Year, month, day, 0, 0, 0, d.Offset);
                    }
                    else {
                        d = new DateTimeOffset(d.Year, month, daysInMonth, 0, 0, 0, d.Offset).AddDays(day - daysInMonth);
                    }
                }

                return new NextFireTimeCursor(true, d);
            }

            return new NextFireTimeCursor(false, d);
        }

        private NextFireTimeCursor ProgressNextFireTimeDayOfWeek(DateTimeOffset d) {
            var day = d.Day;
            var mon = d.Month;

            // get day by day of week rule
            if (this.lastDayOfWeek) {
                // are we looking for the last XXX day of
                // the month?
                var dow = this.daysOfWeek.Min; // desired
                                               // d-o-w
                var cDow = (int)d.DayOfWeek + 1; // current d-o-w
                var daysToAdd = 0;
                if (cDow < dow) {
                    daysToAdd = dow - cDow;
                }

                if (cDow > dow) {
                    daysToAdd = dow + (7 - cDow);
                }

                var lDay = GetLastDayOfMonth(mon, d.Year);

                if (day + daysToAdd > lDay) {
                    // did we already miss the
                    // last one?
                    if (mon == 12) {
                        //will we pass the end of the year?
                        d = new DateTimeOffset(d.Year, mon - 11, 1, 0, 0, 0, d.Offset).AddYears(1);
                    }
                    else {
                        d = new DateTimeOffset(d.Year, mon + 1, 1, 0, 0, 0, d.Offset);
                    }

                    // we are promoting the month
                    return new NextFireTimeCursor(true, d);
                }

                // find date of last occurrence of this day in this month...
                while (day + daysToAdd + 7 <= lDay) {
                    daysToAdd += 7;
                }

                day += daysToAdd;

                if (daysToAdd > 0) {
                    // we are not promoting the month
                    return new NextFireTimeCursor(true, new DateTimeOffset(d.Year, mon, day, 0, 0, 0, d.Offset));
                }
            }
            else if (this.nthdayOfWeek != 0) {
                // are we looking for the Nth XXX day in the month?
                var dow = this.daysOfWeek.Min; // desired
                                               // d-o-w
                var cDow = (int)d.DayOfWeek + 1; // current d-o-w
                var daysToAdd = 0;
                if (cDow < dow) {
                    daysToAdd = dow - cDow;
                }
                else if (cDow > dow) {
                    daysToAdd = dow + (7 - cDow);
                }

                var dayShifted = daysToAdd > 0;

                day += daysToAdd;
                var weekOfMonth = day / 7;
                if (day % 7 > 0) {
                    weekOfMonth++;
                }

                daysToAdd = (this.nthdayOfWeek - weekOfMonth) * 7;
                day += daysToAdd;
                if (daysToAdd < 0 || day > GetLastDayOfMonth(mon, d.Year)) {
                    if (mon == 12) {
                        d = new DateTimeOffset(d.Year, mon - 11, 1, 0, 0, 0, d.Offset).AddYears(1);
                    }
                    else {
                        d = new DateTimeOffset(d.Year, mon + 1, 1, 0, 0, 0, d.Offset);
                    }

                    // we are promoting the month
                    return new NextFireTimeCursor(true, d);
                }

                if (daysToAdd > 0 || dayShifted) {
                    // we are NOT promoting the month
                    return new NextFireTimeCursor(true, new DateTimeOffset(d.Year, mon, day, 0, 0, 0, d.Offset));
                }
            }
            else if (this.everyNthWeek != 0) {
                var cDow = (int)d.DayOfWeek + 1; // current d-o-w
                var dow = this.daysOfWeek.Min; // desired d-o-w
                if (this.daysOfWeek.TryGetMinValueStartingFrom(cDow, out var min)) {
                    dow = min;
                }

                var daysToAdd = 0;
                if (cDow < dow) {
                    daysToAdd = (dow - cDow) + (7 * (this.everyNthWeek - 1));
                }

                if (cDow > dow) {
                    daysToAdd = (dow + (7 - cDow)) + (7 * (this.everyNthWeek - 1));
                }

                if (daysToAdd > 0) {
                    // are we switching days?
                    d = new DateTimeOffset(d.Year, mon, day, 0, 0, 0, d.Offset);
                    d = d.AddDays(daysToAdd);
                    return new NextFireTimeCursor(true, d);
                }
            }
            else {
                var cDow = (int)d.DayOfWeek + 1; // current d-o-w
                var dow = this.daysOfWeek.Min; // desired d-o-w
                if (this.daysOfWeek.TryGetMinValueStartingFrom(cDow, out var min)) {
                    dow = min;
                }

                var daysToAdd = 0;
                if (cDow < dow) {
                    daysToAdd = dow - cDow;
                }

                if (cDow > dow) {
                    daysToAdd = dow + (7 - cDow);
                }

                var lDay = GetLastDayOfMonth(mon, d.Year);

                if (day + daysToAdd > lDay) {
                    // will we pass the end of the month?

                    if (mon == 12) {
                        //will we pass the end of the year?
                        d = new DateTimeOffset(d.Year, mon - 11, 1, 0, 0, 0, d.Offset).AddYears(1);
                    }
                    else {
                        d = new DateTimeOffset(d.Year, mon + 1, 1, 0, 0, 0, d.Offset);
                    }

                    // we are promoting the month
                    return new NextFireTimeCursor(true, d);
                }

                if (daysToAdd > 0) {
                    // are we switching days?
                    return new NextFireTimeCursor(true, new DateTimeOffset(d.Year, mon, day + daysToAdd, 0, 0, 0, d.Offset));
                }
            }

            return new NextFireTimeCursor(false, new DateTimeOffset(d.Year, d.Month, day, d.Hour, d.Minute, d.Second, d.Offset));
        }

        /// <summary>
        /// Progress next fire time day
        /// </summary>
        /// <param name="d">NextFireTimeCheck</param>
        private NextFireTimeCursor ProgressNextFireTimeDay(DateTimeOffset d) {
            var dayOfMSpec = !this.daysOfMonth.Contains(CronExpressionConstants.NoSpec);
            var dayOfWSpec = !this.daysOfWeek.Contains(CronExpressionConstants.NoSpec);
            if (dayOfMSpec && !dayOfWSpec) {
                return this.ProgressNextFireTimeDayOfMonth(d);
            }

            if (dayOfWSpec && !dayOfMSpec) {
                return this.ProgressNextFireTimeDayOfWeek(d);
            }

            var dayOfMonthProgressResult = this.ProgressNextFireTimeDayOfMonth(d);
            var dayOfWeekProgressResult = this.ProgressNextFireTimeDayOfWeek(d);
            if (dayOfMonthProgressResult.RestartLoop && dayOfWeekProgressResult.RestartLoop) {
                return dayOfWeekProgressResult.Date < dayOfMonthProgressResult.Date
                    ? dayOfWeekProgressResult
                    : dayOfMonthProgressResult;
            }

            // only 1 result has value then return it
            if (dayOfWeekProgressResult is { Date: { }, RestartLoop: false })
                return dayOfWeekProgressResult;
            if (dayOfMonthProgressResult is { Date: { }, RestartLoop: false })
                return dayOfMonthProgressResult;

            // both results have value, return earliest
            return dayOfWeekProgressResult.Date!.Value < dayOfMonthProgressResult.Date!.Value
                ? dayOfWeekProgressResult
                : dayOfMonthProgressResult;
        }

        /// <summary>
        /// Progress next fire time Month
        /// </summary>
        /// <param name="d">NextFireTimeCheck</param>
        private NextFireTimeCursor ProgressNextFireTimeMonth(DateTimeOffset d) {
            var mon = d.Month;
            var year = d.Year;
            var t = -1;

            if (this.months.TryGetMinValueStartingFrom(mon, out var min)) {
                t = mon;
                mon = min;
            }
            else {
                mon = this.months.Min;
                year++;
            }

            return mon != t
                ? new NextFireTimeCursor(true, new DateTimeOffset(year, mon, 1, 0, 0, 0, d.Offset))
                : new NextFireTimeCursor(false, new DateTimeOffset(d.Year, mon, d.Day, d.Hour, d.Minute, d.Second, d.Offset));
        }

        private NextFireTimeCursor ProgressNextFireTimeYear(DateTimeOffset d) {
            var year = d.Year;
            int t;
            if (this.years.TryGetMinValueStartingFrom(d.Year, out var min)) {
                t = year;
                year = min;
            }
            else {
                // ran out of years...
                return new NextFireTimeCursor(false, null);
            }

            if (year != t) {
                return new NextFireTimeCursor(true, new DateTimeOffset(year, 1, 1, 0, 0, 0, d.Offset));
            }

            return new NextFireTimeCursor(false, new DateTimeOffset(year, d.Month, d.Day, d.Hour, d.Minute, d.Second, d.Offset));
        }

        /// <summary>
        /// Gets the next fire time after the given time.
        /// </summary>
        /// <param name="afterTimeUtc">The UTC time to start searching from.</param>
        /// <returns></returns>
        public DateTimeOffset? GetTimeAfter(DateTimeOffset afterTimeUtc) {
            // move ahead one second, since we're computing the time *after* the
            // given time
            afterTimeUtc = afterTimeUtc.AddSeconds(1);

            // CronTrigger does not deal with milliseconds
            var d = CreateDateTimeWithoutMilliseconds(afterTimeUtc);

            // change to specified time zone
            d = TimeZoneUtil.ConvertTime(d, this.TimeZone);

            var nextFireTimeProgressors = new[] { this.ProgressNextFireTimeSecond, this.ProgressNextFireTimeMinute, this.ProgressNextFireTimeHour, this.ProgressNextFireTimeDay, this.ProgressNextFireTimeMonth, this.ProgressNextFireTimeYear };

            var nextFireTimeCursor = new NextFireTimeCursor(false, d);
            var foundNextFireTime = false;

            // loop until we've computed the next time, or we've past the endTime
            while (!foundNextFireTime) {
                foreach (var progressor in nextFireTimeProgressors) {
                    if (nextFireTimeCursor.Date.HasValue) {
                        nextFireTimeCursor = progressor(nextFireTimeCursor.Date.Value);
                    }
                    else {
                        break;
                    }

                    if (nextFireTimeCursor.RestartLoop) {
                        break;
                    }
                }

                // test for expressions that never generate a valid fire date,
                if (nextFireTimeCursor.Date is null || nextFireTimeCursor.Date.Value.Year > TriggerConstants.YearToGiveUpSchedulingAt) {
                    return null; // ran out of years
                }

                if (nextFireTimeCursor.RestartLoop) {
                    continue;
                }

                // apply the proper offset for this date
                d = new DateTimeOffset(nextFireTimeCursor.Date.Value.DateTime, TimeZoneUtil.GetUtcOffset(nextFireTimeCursor.Date.Value.DateTime, this.TimeZone));
                foundNextFireTime = true;
            }

            return d.ToUniversalTime();
        }

        /// <summary>
        /// Creates the date time without milliseconds.
        /// </summary>
        /// <param name="time">The time.</param>
        /// <returns></returns>
        private static DateTimeOffset CreateDateTimeWithoutMilliseconds(DateTimeOffset time) {
            return new DateTimeOffset(time.Year, time.Month, time.Day, time.Hour, time.Minute, time.Second, time.Offset);
        }

        /// <summary>
        /// Advance the calendar to the particular hour paying particular attention
        /// to daylight saving problems.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <param name="hour">The hour.</param>
        /// <returns></returns>
        private static DateTimeOffset SetCalendarHour(DateTimeOffset date, int hour) {
            // Java version of Quartz uses lenient calendar
            // so hour 24 creates day increment and zeroes hour
            var hourToSet = hour;
            if (hourToSet == 24) {
                hourToSet = 0;
            }

            var d = new DateTimeOffset(date.Year, date.Month, date.Day, hourToSet, date.Minute, date.Second, date.Millisecond, date.Offset);
            if (hour == 24) {
                // increment day
                d = d.AddDays(1);
            }

            return d;
        }

        /// <summary>
        /// Gets the time before.
        /// </summary>
        /// <param name="endTime">The end time.</param>
        /// <returns></returns>
#pragma warning disable CA1822
        public DateTimeOffset? GetTimeBefore(DateTimeOffset? endTime)
#pragma warning restore CA1822
        {
            // TODO: implement
            return null;
        }

        /// <summary>
        /// NOT YET IMPLEMENTED: Returns the final time that the
        /// <see cref="CronExpression" /> will match.
        /// </summary>
        /// <returns></returns>
#pragma warning disable CA1822
        public DateTimeOffset? GetFinalFireTime()
#pragma warning restore CA1822
        {
            // TODO: implement QUARTZ-423
            return null;
        }

        /// <summary>
        /// Gets the last day of month.
        /// </summary>
        private static int GetLastDayOfMonth(int monthNum, int year) {
            return DateTime.DaysInMonth(year, monthNum);
        }

        private static int ToInt32(char c) => c - '0';

        private static int ToInt32(ReadOnlySpan<char> span) {
            return int.Parse(span);
        }

        /// <summary>
        /// Creates a new object that is a copy of the current instance.
        /// </summary>
        /// <returns>
        /// A new object that is a copy of this instance.
        /// </returns>
        public object Clone() {
            return new CronExpression(this.CronExpressionString) { TimeZone = this.TimeZone };
        }

        /// <summary>
        /// Determines whether the specified <see cref="CronExpression"/> is equal to the current <see cref="CronExpression"/>.
        /// </summary>
        /// <returns>
        /// true if the specified <see cref="CronExpression"/> is equal to the current <see cref="CronExpression"/>; otherwise, false.
        /// </returns>
        /// <param name="other">The <see cref="CronExpression"/> to compare with the current <see cref="CronExpression"/>. </param>
        public bool Equals(CronExpression other) {
            if (other is null) return false;
            if (ReferenceEquals(this, other)) return true;
            return Equals(other.CronExpressionString, this.CronExpressionString) && Equals(other.TimeZone, this.TimeZone);
        }

        /// <summary>
        /// Determines whether the specified <see cref="System.Object"/> is equal to the current <see cref="System.Object"/>.
        /// </summary>
        /// <returns>
        /// true if the specified <see cref="System.Object"/> is equal to the current <see cref="System.Object"/>; otherwise, false.
        /// </returns>
        /// <param name="obj">The <see cref="System.Object"/> to compare with the current <see cref="System.Object"/>. </param>
        public override bool Equals(object? obj) {
            if (obj is null) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(CronExpression)) return false;
            return this.Equals((CronExpression)obj);
        }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="System.Object"/>.
        /// </returns>
        /// <filterpriority>2</filterpriority>
        public override int GetHashCode() {
            unchecked {
                return ((this.CronExpressionString is not null ? this.CronExpressionString.GetHashCode() : 0) * 397) ^ (this.timeZone is not null ? this.timeZone.GetHashCode() : 0);
            }
        }
    }

    [StructLayout(LayoutKind.Auto)]
    internal readonly record struct ValueAndPosition(int Value, int Position);

    /// <summary>
    /// </summary>
    /// <param name="RestartLoop">Indicate if the Next fire date progressor loop should restart</param>
    /// <param name="Date">NextFireDate calculated progress result</param>
    [StructLayout(LayoutKind.Auto)]
    internal readonly record struct NextFireTimeCursor(bool RestartLoop, DateTimeOffset? Date);

    /// <summary>
    /// Optimized structure to hold either one value or multiple.
    /// </summary>
    internal sealed class CronField : IEnumerable<int> {
        // null == not set, all spec or individual value
        private int? singleValue;
        private SortedSet<int>? values;
        private bool hasAllOrNoSpec;

        public CronField() {
            this.Clear();
        }

        internal int Count {
            get {
                if (this.singleValue is not null) {
                    return 1;
                }

                return this.values?.Count ?? 0;
            }
        }

        internal int Min {
            get {
                if (this.singleValue is not null) {
                    return this.hasAllOrNoSpec ? 0 : this.singleValue.Value;
                }

                if (this.values is not null) {
                    return this.hasAllOrNoSpec ? 0 : this.values.Min;
                }

                return 0;
            }
        }

        internal void Clear() {
            this.singleValue = null;
            this.values = null;
            this.hasAllOrNoSpec = false;
        }

        internal bool TryGetMinValueStartingFrom(int start, out int min) {
            min = 0;

            if (this.singleValue == CronExpressionConstants.AllSpec) {
                min = start;
                return true;
            }

            if (this.singleValue is not null) {
                if (this.singleValue >= start) {
                    min = this.singleValue.Value;
                    return true;
                }

                // didn't match
                return false;
            }

            var set = this.values;

            if (set is null) {
                return false;
            }

            min = set.Min;

            if (set.Contains(start)) {
                min = start;
                return true;
            }

            if (set.Count == 0 || set.Max < start) {
                return false;
            }

            if (set.Min >= start) {
                // value is contained and would be returned from view
                return true;
            }

            // slow path
            var view = set.GetViewBetween(start, int.MaxValue);
            if (view.Count > 0) {
                min = view.Min;
                return true;
            }

            return false;
        }

        public void Add(int value) {
            this.hasAllOrNoSpec = value is CronExpressionConstants.AllSpec or CronExpressionConstants.NoSpec;

            if (this.singleValue is null) {
                if (this.values is null) {
                    this.singleValue = value;
                }
                else {
                    this.values.Add(value);
                }
            }
            else if (this.singleValue != value) {
                this.values = new SortedSet<int> { this.singleValue.Value, value };
                this.singleValue = null;
            }
        }

        public bool Contains(int value) {
            if (this.singleValue == value
                || (value != CronExpressionConstants.AllSpec && value != CronExpressionConstants.NoSpec && this.hasAllOrNoSpec)) {
                return true;
            }

            return this.values is not null && this.values.Contains(value);
        }

        public IEnumerator<int> GetEnumerator() {
            if (this.singleValue is not null) {
                yield return this.singleValue.Value;
                yield break;
            }

            if (this.values is not null) {
                foreach (var value in this.values) {
                    yield return value;
                }
            }
        }

        IEnumerator IEnumerable.GetEnumerator() {
            return this.GetEnumerator();
        }
    }

    public static class TimeZoneUtil {
        private static readonly Dictionary<string, string> timeZoneIdAliases = new Dictionary<string, string>();

        static TimeZoneUtil() {
            // Azure has had issues with having both formats
            timeZoneIdAliases["UTC"] = "Coordinated Universal Time";
            timeZoneIdAliases["Coordinated Universal Time"] = "UTC";

            // Mono differs in naming too...
            timeZoneIdAliases["Central European Standard Time"] = "CET";
            timeZoneIdAliases["CET"] = "Central European Standard Time";

            timeZoneIdAliases["Eastern Standard Time"] = "US/Eastern";
            timeZoneIdAliases["US/Eastern"] = "Eastern Standard Time";

            timeZoneIdAliases["Central Standard Time"] = "US/Central";
            timeZoneIdAliases["US/Central"] = "Central Standard Time";

            timeZoneIdAliases["US Central Standard Time"] = "US/Indiana-Stark";
            timeZoneIdAliases["US/Indiana-Stark"] = "US Central Standard Time";

            timeZoneIdAliases["Mountain Standard Time"] = "US/Mountain";
            timeZoneIdAliases["US/Mountain"] = "Mountain Standard Time";

            timeZoneIdAliases["US Mountain Standard Time"] = "US/Arizona";
            timeZoneIdAliases["US/Arizona"] = "US Mountain Standard Time";

            timeZoneIdAliases["Pacific Standard Time"] = "US/Pacific";
            timeZoneIdAliases["US/Pacific"] = "Pacific Standard Time";

            timeZoneIdAliases["Alaskan Standard Time"] = "US/Alaska";
            timeZoneIdAliases["US/Alaska"] = "Alaskan Standard Time";

            timeZoneIdAliases["Hawaiian Standard Time"] = "US/Hawaii";
            timeZoneIdAliases["US/Hawaii"] = "Hawaiian Standard Time";

            timeZoneIdAliases["China Standard Time"] = "Asia/Shanghai";
            timeZoneIdAliases["Asia/Shanghai"] = "China Standard Time";

            timeZoneIdAliases["Pakistan Standard Time"] = "Asia/Karachi";
            timeZoneIdAliases["Asia/Karachi"] = "Pakistan Standard Time";
        }

#pragma warning disable CA2211
        public static Func<string, TimeZoneInfo?> CustomResolver = id => null;
#pragma warning restore CA2211

        /// <summary>
        /// TimeZoneInfo.ConvertTime is not supported under mono
        /// </summary>
        /// <param name="dateTimeOffset"></param>
        /// <param name="timeZoneInfo"></param>
        /// <returns></returns>
        public static DateTimeOffset ConvertTime(DateTimeOffset dateTimeOffset, TimeZoneInfo timeZoneInfo) {
            return TimeZoneInfo.ConvertTime(dateTimeOffset, timeZoneInfo);
        }

        /// <summary>
        /// TimeZoneInfo.GetUtcOffset(DateTimeOffset) is not supported under mono
        /// </summary>
        /// <param name="dateTimeOffset"></param>
        /// <param name="timeZoneInfo"></param>
        /// <returns></returns>
        public static TimeSpan GetUtcOffset(DateTimeOffset dateTimeOffset, TimeZoneInfo timeZoneInfo) {

            return timeZoneInfo.GetUtcOffset(dateTimeOffset);
        }

        public static TimeSpan GetUtcOffset(DateTime dateTime, TimeZoneInfo timeZoneInfo) {
            // Unlike the default behavior of TimeZoneInfo.GetUtcOffset, it is prefered to choose
            // the DAYLIGHT time when the input is ambiguous, because the daylight instance is the
            // FIRST instance, and time moves in a forward direction.

            var offset = timeZoneInfo.IsAmbiguousTime(dateTime)
                ? timeZoneInfo.GetAmbiguousTimeOffsets(dateTime).Max()
                : timeZoneInfo.GetUtcOffset(dateTime);

            return offset;
        }

        /// <summary>
        /// Tries to find time zone with given id, has ability do some fallbacks when necessary.
        /// </summary>
        /// <param name="id">System id of the time zone.</param>
        /// <returns></returns>
        public static TimeZoneInfo FindTimeZoneById(string id) {
            TimeZoneInfo? info = null;
            try {
                info = TimeZoneInfo.FindSystemTimeZoneById(id);
            }
            catch (TimeZoneNotFoundException ex) {
                if (timeZoneIdAliases.TryGetValue(id, out var aliasedId)) {
                    try {
                        info = TimeZoneInfo.FindSystemTimeZoneById(aliasedId);
                    }
                    catch {
                        Console.WriteLine("Could not find time zone using alias id {AliasId}", aliasedId);
                    }
                }

                if (info is null) {
                    info = CustomResolver(id);
                }

                if (info is null) {
                    // we tried our best
                    throw new TimeZoneNotFoundException(
                        $"Could not find time zone with id {id}, consider using Quartz.Plugins.TimeZoneConverter for resolving more time zones ids",
                        ex);
                }
            }

            return info;
        }
    }

    internal static class CronExpressionConstants {
        /// <summary>
        /// Field specification for second.
        /// </summary>
        public const int Second = 0;

        /// <summary>
        /// Field specification for minute.
        /// </summary>
        public const int Minute = 1;

        /// <summary>
        /// Field specification for hour.
        /// </summary>
        public const int Hour = 2;

        /// <summary>
        /// Field specification for day of month.
        /// </summary>
        public const int DayOfMonth = 3;

        /// <summary>
        /// Field specification for month.
        /// </summary>
        public const int Month = 4;

        /// <summary>
        /// Field specification for day of week.
        /// </summary>
        public const int DayOfWeek = 5;

        /// <summary>
        /// Field specification for year.
        /// </summary>
        public const int Year = 6;

        /// <summary>
        /// Field specification for wildcard '*'.
        /// </summary>
        public const int AllSpec = 99;

        /// <summary>
        /// Field specification for no specification at all '?'.
        /// </summary>
        public const int NoSpec = 98;
    }

    /// <summary>
    /// Extension methods for <see cref="string" />.
    /// </summary>
    internal static class StringExtensions {
        /// <summary>
        /// Allows null-safe trimming of string.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        internal static string? NullSafeTrim(this string? s) {
            return s?.Trim();
        }

        /// <summary>
        /// Trims string and if resulting string is empty, null is returned.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        internal static string? TrimEmptyToNull(this string s) {
            if (s is null) {
                return null;
            }

            s = s.Trim();

            if (s.Length == 0) {
                return null;
            }

            return s;
        }

        // based on https://www.meziantou.net/split-a-string-into-lines-without-allocation.htm
        internal static StringSplitEnumerator SpanSplit(this string str, char ch1, char ch2 = char.MinValue) => SpanSplit(str.AsSpan(), ch1, ch2);

        internal static StringSplitEnumerator SpanSplit(this ReadOnlySpan<char> span, char ch1, char ch2 = char.MinValue) => new(span, ch1, ch2);

        // Must be a ref struct as it contains a ReadOnlySpan<char>
        [StructLayout(LayoutKind.Auto)]
        internal ref struct StringSplitEnumerator {
            private ReadOnlySpan<char> _str;
            private readonly char ch1;
            private readonly char ch2;

            public StringSplitEnumerator(ReadOnlySpan<char> str, char ch1, char ch2) {
                this._str = str;
                this.ch1 = ch1;
                this.ch2 = ch2;
                this.Current = default;
            }

            // Needed to be compatible with the foreach operator
            public StringSplitEnumerator GetEnumerator() => this;

            public bool MoveNext() {
                var span = this._str;
                if (span.Length == 0) // Reach the end of the string
                    return false;

                var index = this.ch2 != char.MinValue
                    ? span.IndexOfAny(this.ch1, this.ch2)
                    : span.IndexOf(this.ch1);

                if (index == -1) // The string is composed of only token
                {
                    this._str = ReadOnlySpan<char>.Empty; // The remaining string is an empty string
                    this.Current = new StringSplitEntry(span, ReadOnlySpan<char>.Empty);
                    return true;
                }

                this.Current = new StringSplitEntry(span.Slice(0, index), span.Slice(index, 1));
                this._str = span.Slice(index + 1);
                return true;
            }

            public StringSplitEntry Current { get; private set; }
        }

        [StructLayout(LayoutKind.Auto)]
        internal readonly ref struct StringSplitEntry {
            public StringSplitEntry(ReadOnlySpan<char> token, ReadOnlySpan<char> separator) {
                this.Token = token;
                this.Separator = separator;
            }

            public ReadOnlySpan<char> Token { get; }
            public ReadOnlySpan<char> Separator { get; }

            // This method allow to deconstruct the type, so you can write any of the following code
            // foreach (var entry in str.SplitLines()) { _ = entry.Line; }
            // foreach (var (line, endOfLine) in str.SplitLines()) { _ = line; }
            // https://learn.microsoft.com/en-us/dotnet/csharp/fundamentals/functional/deconstruct?WT.mc_id=DT-MVP-5003978#deconstructing-user-defined-types
            public void Deconstruct(out ReadOnlySpan<char> line, out ReadOnlySpan<char> separator) {
                line = this.Token;
                separator = this.Separator;
            }

            // This method allow to implicitly cast the type into a ReadOnlySpan<char>, so you can write the following code
            // foreach (ReadOnlySpan<char> entry in str.SplitLines())
            public static implicit operator ReadOnlySpan<char>(StringSplitEntry entry) => entry.Token;
        }
    }
    internal readonly struct CronExpressionSummary {
        public CronExpressionSummary(CronField seconds, CronField minutes, CronField hours, CronField daysOfMonth,
            CronField months, CronField daysOfWeek, bool lastDayOfWeek, bool nearestWeekday, int nthDayOfWeek,
            bool lastDayOfMonth, bool calendarDayOfWeek, bool calendarDayOfMonth, CronField years) {

            this.Seconds = seconds;
            this.Minutes = minutes;
            this.Hours = hours;
            this.DaysOfMonth = daysOfMonth;
            this.Months = months;
            this.DaysOfWeek = daysOfWeek;
            this.LastDayOfWeek = lastDayOfWeek;
            this.NearestWeekday = nearestWeekday;
            this.NthDayOfWeek = nthDayOfWeek;
            this.LastDayOfMonth = lastDayOfMonth;
            this.CalendarDayOfWeek = calendarDayOfWeek;
            this.CalendarDayOfMonth = calendarDayOfMonth;
            this.Years = years;
        }

        public CronField Seconds { get; }
        public CronField Minutes { get; }
        public CronField Hours { get; }
        public CronField DaysOfMonth { get; }
        public CronField Months { get; }
        public CronField DaysOfWeek { get; }
        public bool LastDayOfWeek { get; }
        public bool NearestWeekday { get; }
        public int NthDayOfWeek { get; }
        public bool LastDayOfMonth { get; }
        public bool CalendarDayOfWeek { get; }
        public bool CalendarDayOfMonth { get; }
        public CronField Years { get; }

        /// <summary>
        /// Gets the expression set summary.
        /// </summary>
        private static string GetExpressionSetSummary(CronField data) {

            if (data.Contains(CronExpressionConstants.NoSpec)) {
                return "?";
            }

            if (data.Contains(CronExpressionConstants.AllSpec)) {
                return "*";
            }

            var buf = new StringBuilder();

            var first = true;
            foreach (var iVal in data) {
                var val = iVal.ToString(CultureInfo.InvariantCulture);
                if (!first) {
                    buf.Append(',');
                }

                buf.Append(val);
                first = false;
            }

            return buf.ToString();
        }

        public override string ToString() {

            var buf = new StringBuilder();

            buf.Append("seconds: ");
            buf.AppendLine(GetExpressionSetSummary(this.Seconds));
            buf.Append("minutes: ");
            buf.AppendLine(GetExpressionSetSummary(this.Minutes));
            buf.Append("hours: ");
            buf.AppendLine(GetExpressionSetSummary(this.Hours));
            buf.Append("daysOfMonth: ");
            buf.AppendLine(GetExpressionSetSummary(this.DaysOfMonth));
            buf.Append("months: ");
            buf.AppendLine(GetExpressionSetSummary(this.Months));
            buf.Append("daysOfWeek: ");
            buf.AppendLine(GetExpressionSetSummary(this.DaysOfWeek));
            buf.Append("lastdayOfWeek: ");
            buf.AppendLine(this.LastDayOfWeek.ToString());
            buf.Append("nearestWeekday: ");
            buf.AppendLine(this.NearestWeekday.ToString());
            buf.Append("NthDayOfWeek: ");
            buf.AppendLine(this.NthDayOfWeek.ToString());
            buf.Append("lastdayOfMonth: ");
            buf.AppendLine(this.LastDayOfMonth.ToString());
            buf.Append("calendardayOfWeek: ");
            buf.AppendLine(this.CalendarDayOfWeek.ToString());
            buf.Append("calendardayOfMonth: ");
            buf.AppendLine(this.CalendarDayOfMonth.ToString());
            buf.Append("years: ");
            buf.AppendLine(GetExpressionSetSummary(this.Years));
            return buf.ToString();
        }
    }
    public static class TriggerConstants {
        /// <summary>
        /// The default value for priority.
        /// </summary>
        public const int DefaultPriority = 5;

        internal static readonly int YearToGiveUpSchedulingAt = TimeProvider.System.GetUtcNow().Year + 100;
        internal const int EarliestYear = 1970;
    }

    internal static class SortedSetExtensions {
        internal static bool TryGetMinValueStartingFrom(this SortedSet<int> set, DateTimeOffset start, bool allowValueBeforeStartDay, out int minimumDay) {

            minimumDay = set.Min;
            var startDay = start.Day;

            if (set.Contains(CronExpressionConstants.AllSpec) || set.Contains(startDay)) {
                minimumDay = startDay;
                return true;
            }

            // In cases such as W modifier finding a match earlier than the month day.
            // If the flag allowValueBeforeStartDay is set and the minimum value is less than the start day, return the minimum value 
            if (allowValueBeforeStartDay && set.Min < startDay) {
                return true;
            }

            // If the set is empty or the maximum value is less than the start day, no suitable value is found
            if (set.Count == 0 || set.Max < startDay) {
                return false;
            }

            // If the minimum value is greater than or equal to the start day, return the minimum value
            if (set.Min >= startDay) {
                return true;
            }

            // slow path
            var view = set.GetViewBetween(startDay, int.MaxValue);
            if (view.Count > 0) {
                minimumDay = view.Min;
                return true;
            }

            return false;
        }
    }
}
