using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace me.fengyj.CommonLib.Utils.Scheduler {
    internal class CronExpressionDescriptor {
        private readonly char[] m_specialCharacters = ['/', '-', ',', '*'];

        private string m_expression;
        private Options m_options;
        private string[] m_expressionParts;
        private bool m_parsed;
        private bool m_use24HourTimeFormat;
        private CultureInfo m_culture;

        /// <summary>
        /// Initializes a new instance of the <see cref="CronExpressionDescriptor"/> class
        /// </summary>
        /// <param name="expression">The cron expression string</param>
        internal CronExpressionDescriptor(string expression) : this(expression, new Options()) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="CronExpressionDescriptor"/> class
        /// </summary>
        /// <param name="expression">The cron expression string</param>
        /// <param name="options">Options to control the output description</param>
        internal CronExpressionDescriptor(string expression, Options options) {
            this.m_expression = expression;
            this.m_options = options;
            this.m_expressionParts = new string[7];
            this.m_parsed = false;

            if (!string.IsNullOrEmpty(options.Locale)) {
                this.m_culture = new CultureInfo(options.Locale);
            }
            else {
                // If options.Locale not specified...
#if NET_STANDARD_1X
        // .NET Standard 1.* will use English as default
        m_culture = new CultureInfo("en-US");
#else
                // .NET Standard >= 2.0 will use CurrentUICulture as default
                this.m_culture = System.Threading.Thread.CurrentThread.CurrentUICulture;
#endif
            }

            if (this.m_options.Use24HourTimeFormat != null) {
                // 24HourTimeFormat specified in options so use it
                this.m_use24HourTimeFormat = this.m_options.Use24HourTimeFormat.Value;
            }
            else {
                // 24HourTimeFormat not specified, default based on m_24hourTimeFormatLocales
                this.m_use24HourTimeFormat = true;
            }
        }

        /// <summary>
        /// Generates a human readable string for the Cron Expression
        /// </summary>
        /// <param name="type">Which part(s) of the expression to describe</param>
        /// <returns>The cron expression description</returns>
        public string? GetDescription(DescriptionTypeEnum type) {
            var description = string.Empty;

            try {
                if (!this.m_parsed) {
                    var parser = new ExpressionParser(this.m_expression, this.m_options);
                    this.m_expressionParts = parser.Parse();
                    this.m_parsed = true;
                }

                switch (type) {
                    case DescriptionTypeEnum.FULL:
                        description = this.GetFullDescription();
                        break;
                    case DescriptionTypeEnum.TIMEOFDAY:
                        description = this.GetTimeOfDayDescription();
                        break;
                    case DescriptionTypeEnum.HOURS:
                        description = this.GetHoursDescription();
                        break;
                    case DescriptionTypeEnum.MINUTES:
                        description = this.GetMinutesDescription();
                        break;
                    case DescriptionTypeEnum.SECONDS:
                        description = this.GetSecondsDescription();
                        break;
                    case DescriptionTypeEnum.DAYOFMONTH:
                        description = this.GetDayOfMonthDescription();
                        break;
                    case DescriptionTypeEnum.MONTH:
                        description = this.GetMonthDescription();
                        break;
                    case DescriptionTypeEnum.DAYOFWEEK:
                        description = this.GetDayOfWeekDescription();
                        break;
                    case DescriptionTypeEnum.YEAR:
                        description = this.GetYearDescription();
                        break;
                    default:
                        description = this.GetSecondsDescription();
                        break;
                }
            }
            catch (Exception ex) {
                if (!this.m_options.ThrowExceptionOnParseError) {
                    description = ex.Message;
                }
                else {
                    throw;
                }
            }

            // Uppercase the first letter
            if (description != null)
                description = string.Concat(this.m_culture.TextInfo.ToUpper(description[0]), description.Substring(1));

            return description;
        }

        /// <summary>
        /// Generates the FULL description
        /// </summary>
        /// <returns>The FULL description</returns>
        protected string GetFullDescription() {
            string description;

            try {
                var timeSegment = this.GetTimeOfDayDescription();
                var dayOfMonthDesc = this.GetDayOfMonthDescription();
                var monthDesc = this.GetMonthDescription();
                var dayOfWeekDesc = this.GetDayOfWeekDescription();
                var yearDesc = this.GetYearDescription();

                description = string.Format("{0}{1}{2}{3}{4}",
                       timeSegment,
                       dayOfMonthDesc,
                       dayOfWeekDesc,
                       monthDesc,
                       yearDesc);

                description = this.TransformVerbosity(description, this.m_options.Verbose);
            }
            catch (Exception ex) {
                description = "An error occured when generating the expression description.  Check the cron expression syntax.";
                if (this.m_options.ThrowExceptionOnParseError) {
                    throw new FormatException(description, ex);
                }
            }


            return description;
        }

        /// <summary>
        /// Generates a description for only the TIMEOFDAY portion of the expression
        /// </summary>
        /// <returns>The TIMEOFDAY description</returns>
        protected string GetTimeOfDayDescription() {
            var secondsExpression = this.m_expressionParts[0];
            var minuteExpression = this.m_expressionParts[1];
            var hourExpression = this.m_expressionParts[2];

            var description = new StringBuilder();

            //handle special cases first
            if (minuteExpression.IndexOfAny(this.m_specialCharacters) == -1
                && hourExpression.IndexOfAny(this.m_specialCharacters) == -1
                && secondsExpression.IndexOfAny(this.m_specialCharacters) == -1) {
                //specific time of day (i.e. 10 14)
                description.Append("At ").Append(this.FormatTime(hourExpression, minuteExpression, secondsExpression));
            }
            else if (secondsExpression == "" && minuteExpression.Contains("-")
                && !minuteExpression.Contains(",")
                && hourExpression.IndexOfAny(this.m_specialCharacters) == -1) {
                //minute range in single hour (i.e. 0-10 11)
                var minuteParts = minuteExpression.Split('-');
                description.Append(string.Format("Every minute between {0} and {1}",
                    this.FormatTime(hourExpression, minuteParts[0]),
                    this.FormatTime(hourExpression, minuteParts[1])));
            }
            else if (secondsExpression == "" && hourExpression.Contains(",")
                && hourExpression.IndexOf('-') == -1
                && minuteExpression.IndexOfAny(this.m_specialCharacters) == -1) {
                //hours list with single minute (o.e. 30 6,14,16)
                var hourParts = hourExpression.Split(',');
                description.Append("At");
                for (var i = 0; i < hourParts.Length; i++) {
                    description.Append(" ").Append(this.FormatTime(hourParts[i], minuteExpression));

                    if (i < (hourParts.Length - 2)) {
                        description.Append(",");
                    }

                    if (i == hourParts.Length - 2) {
                        description.Append(" and");
                    }
                }
            }
            else {
                //default time description
                var secondsDescription = this.GetSecondsDescription();
                var minutesDescription = this.GetMinutesDescription();
                var hoursDescription = this.GetHoursDescription();

                description.Append(secondsDescription);

                if (description.Length > 0 && (minutesDescription?.Length ?? 0) > 0) {
                    description.Append(", ");
                }

                description.Append(minutesDescription);

                if (description.Length > 0 && hourExpression.Length > 0) {
                    description.Append(", ");
                }

                description.Append(hoursDescription);
            }


            return description.ToString();
        }

        /// <summary>
        /// Generates a description for only the SECONDS portion of the expression
        /// </summary>
        /// <returns>The SECONDS description</returns>
        protected string? GetSecondsDescription() {
            var description = this.GetSegmentDescription(
               this.m_expressionParts[0],
               "every second",
               (s => s),
               (s => string.Format("every {0} seconds", s)),
               (s => "seconds {0} through {1} past the minute"),
               (s => {
                   if (int.TryParse(s, out var i)) {
                       return s == "0"
                        ? string.Empty
                        : (i < 20)
                            ? "at {0} seconds past the minute"
                            : "at {0} seconds past the minute";

                   }
                   else {
                       return "at {0} seconds past the minute";
                   }
               }),
               (s => ", {0} through {1}")
               );

            return description;
        }

        /// <summary>
        /// Generates a description for only the MINUTE portion of the expression
        /// </summary>
        /// <returns>The MINUTE description</returns>
        protected string? GetMinutesDescription() {
            var secondsExpression = this.m_expressionParts[0];
            var description = this.GetSegmentDescription(
                expression: this.m_expressionParts[1],
                allDescription: "every minute",
                getSingleItemDescription: (s => s),
                getIntervalDescriptionFormat: (s => string.Format("every {0} minutes", s)),
                getBetweenDescriptionFormat: (s => "minutes {0} through {1} past the hour"),
                getDescriptionFormat: (s => {
                    if (int.TryParse(s, out var i)) {
                        return s == "0" && secondsExpression == ""
                          ? string.Empty
                          : (int.Parse(s) < 20)
                              ? "at {0} minutes past the hour"
                              : "at {0} minutes past the hour";
                    }
                    else {
                        return "at {0} minutes past the hour";
                    }
                }),
                getRangeFormat: (s => ", {0} through {1}")
            );

            return description;
        }

        /// <summary>
        /// Generates a description for only the HOUR portion of the expression
        /// </summary>
        /// <returns>The HOUR description</returns>
        protected string? GetHoursDescription() {
            var expression = this.m_expressionParts[2];
            var description = this.GetSegmentDescription(expression,
                   "every hour",
                   (s => this.FormatTime(s, "0")),
                   (s => string.Format("every {0} hours", s)),
                   (s => "between {0} and {1}"),
                   (s => "at {0}"),
                   (s => ", {0} through {1}")
               );

            return description;
        }

        /// <summary>
        /// Generates a description for only the DAYOFWEEK portion of the expression
        /// </summary>
        /// <returns>The DAYOFWEEK description</returns>
        protected string? GetDayOfWeekDescription() {
            string? description = null;

            if (this.m_expressionParts[5] == "*") {
                // DOW is specified as * so we will not generate a description and defer to DOM part.
                // Otherwise, we could get a contradiction like "on day 1 of the month, every day"
                // or a dupe description like "every day, every day".
                description = string.Empty;

            }
            else {
                description = this.GetSegmentDescription(
                    this.m_expressionParts[5],
                    ", every day",
                    (s => {
                        var exp = s.Contains("#")
                            ? s.Remove(s.IndexOf("#"))
                            : s.Contains("L")
                                ? s.Replace("L", string.Empty)
                                : s;

                        return this.m_culture.DateTimeFormat.GetDayName(((DayOfWeek)Convert.ToInt32(exp)));
                    }),
                    (s => string.Format(", every {0} days of the week", s)),
                    (s => ", {0} through {1}"),
                    (s => {
                        string? format = null;
                        if (s.Contains("#")) {
                            var dayOfWeekOfMonthNumber = s.Substring(s.IndexOf("#") + 1);
                            string? dayOfWeekOfMonthDescription = null;
                            switch (dayOfWeekOfMonthNumber) {
                                case "1":
                                    dayOfWeekOfMonthDescription = "first";
                                    break;
                                case "2":
                                    dayOfWeekOfMonthDescription = "second";
                                    break;
                                case "3":
                                    dayOfWeekOfMonthDescription = "third";
                                    break;
                                case "4":
                                    dayOfWeekOfMonthDescription = "fourth";
                                    break;
                                case "5":
                                    dayOfWeekOfMonthDescription = "fifth";
                                    break;
                            }

                            format = string.Concat(", on the ",
                                dayOfWeekOfMonthDescription, " {0} of the month");
                        }
                        else if (s.Contains("L")) {
                            format = ", on the last {0} of the month";
                        }
                        else {
                            format = ", only on {0}";
                        }

                        return format;
                    }),
                    (s => ", {0} through {1}")
              );
            }

            return description;
        }

        /// <summary>
        /// Generates a description for only the MONTH portion of the expression
        /// </summary>
        /// <returns>The MONTH description</returns>
        protected string? GetMonthDescription() {
            var description = this.GetSegmentDescription(
                this.m_expressionParts[4],
                string.Empty,
               (s => new DateTime(DateTime.Now.Year, Convert.ToInt32(s), 1).ToString("MMMM", this.m_culture)),
               (s => string.Format(", every {0} months", s)),
               (s => ", {0} through {1}"),
               (s => ", only in {0}"),
               (s => ", {0} through {1}")
            );

            return description;
        }

        /// <summary>
        /// Generates a description for only the DAYOFMONTH portion of the expression
        /// </summary>
        /// <returns>The DAYOFMONTH description</returns>
        protected string? GetDayOfMonthDescription() {
            string? description = null;
            var expression = this.m_expressionParts[3];

            switch (expression) {
                case "L":
                    description = ", on the last day of the month";
                    break;
                case "WL":
                case "LW":
                    description = ", on the last weekday of the month";
                    break;
                default:
                    var weekDayNumberMatches = new Regex("(\\d{1,2}W)|(W\\d{1,2})");
                    if (weekDayNumberMatches.IsMatch(expression)) {
                        var m = weekDayNumberMatches.Match(expression);
                        var dayNumber = Int32.Parse(m.Value.Replace("W", ""));

                        var dayString = dayNumber == 1 ? "first weekday" :
                            String.Format("weekday nearest day {0}", dayNumber);
                        description = String.Format(", on the {0} of the month", dayString);

                        break;
                    }
                    else {
                        // Handle "last day offset" (i.e. L-5:  "5 days before the last day of the month")
                        var lastDayOffSetMatches = new Regex("L-(\\d{1,2})");
                        if (lastDayOffSetMatches.IsMatch(expression)) {
                            var m = lastDayOffSetMatches.Match(expression);
                            var offSetDays = m.Groups[1].Value;
                            description = String.Format(", {0} days before the last day of the month", offSetDays);
                            break;
                        }
                        else {
                            description = this.GetSegmentDescription(expression,
                                ", every day",
                                (s => s),
                                (s => s == "1" ? ", every day" : ", every {0} days"),
                                (s => ", between day {0} and {1} of the month"),
                                (s => ", on day {0} of the month"),
                                (s => ", {0} through {1}")
                            );
                            break;
                        }
                    }
            }

            return description;
        }

        /// <summary>
        /// Generates a description for only the YEAR portion of the expression
        /// </summary>
        /// <returns>The YEAR description</returns>
        private string? GetYearDescription() {
            var description = this.GetSegmentDescription(this.m_expressionParts[6],
                string.Empty,
               (s => Regex.IsMatch(s, @"^\d+$") ?
                new DateTime(Convert.ToInt32(s), 1, 1).ToString("yyyy") : s),
               (s => string.Format(", every {0} years", s)),
               (s => ", {0} through {1}"),
               (s => ", only in {0}"),
               (s => ", {0} through {1}")
            );

            return description;
        }

        /// <summary>
        /// Generates the segment description
        /// <remarks>
        /// Range expressions used the 'ComaX0ThroughX1' resource
        /// However Romanian language has different idioms for
        /// 1. 'from number to number' (minutes, seconds, hours, days) => ComaMinX0ThroughMinX1 optional resource
        /// 2. 'from month to month' ComaMonthX0ThroughMonthX1 optional resource
        /// 3. 'from year to year' => ComaYearX0ThroughYearX1 optional resource
        /// therefore <paramref name="getRangeFormat"/> was introduced
        /// </remarks>
        /// </summary>
        /// <param name="expression"></param>
        /// <param name="allDescription"></param>
        /// <param name="getSingleItemDescription"></param>
        /// <param name="getIntervalDescriptionFormat"></param>
        /// <param name="getBetweenDescriptionFormat"></param>
        /// <param name="getDescriptionFormat"></param>
        /// <param name="getRangeFormat">function that formats range expressions depending on cron parts</param>
        /// <returns></returns>
        protected string? GetSegmentDescription(string expression,
            string allDescription,
            Func<string, string> getSingleItemDescription,
            Func<string, string> getIntervalDescriptionFormat,
            Func<string, string> getBetweenDescriptionFormat,
            Func<string, string> getDescriptionFormat,
            Func<string, string> getRangeFormat
            ) {
            string? description = null;

            if (string.IsNullOrEmpty(expression)) {
                description = string.Empty;
            }
            else if (expression == "*") {
                description = allDescription;
            }
            else if (expression.IndexOfAny(new char[] { '/', '-', ',' }) == -1) {
                description = string.Format(getDescriptionFormat(expression), getSingleItemDescription(expression));
            }
            else if (expression.Contains("/")) {
                var segments = expression.Split('/');
                description = string.Format(getIntervalDescriptionFormat(segments[1]), getSingleItemDescription(segments[1]));

                //interval contains 'between' piece (i.e. 2-59/3 )
                if (segments[0].Contains("-")) {
                    var betweenSegmentDescription = this.GenerateBetweenSegmentDescription(segments[0], getBetweenDescriptionFormat, getSingleItemDescription);

                    if (!betweenSegmentDescription.StartsWith(", ")) {
                        description += ", ";
                    }

                    description += betweenSegmentDescription;
                }
                else if (segments[0].IndexOfAny(new char[] { '*', ',' }) == -1) {
                    var rangeItemDescription = string.Format(getDescriptionFormat(segments[0]), getSingleItemDescription(segments[0]));
                    //remove any leading comma
                    rangeItemDescription = rangeItemDescription.Replace(", ", "");

                    description += string.Format(", starting {0}", rangeItemDescription);
                }
            }
            else if (expression.Contains(",")) {
                var segments = expression.Split(',');

                var descriptionContent = string.Empty;
                for (var i = 0; i < segments.Length; i++) {
                    if (i > 0 && segments.Length > 2) {
                        descriptionContent += ",";

                        if (i < segments.Length - 1) {
                            descriptionContent += " ";
                        }
                    }

                    if (i > 0 && segments.Length > 1 && (i == segments.Length - 1 || segments.Length == 2)) {
                        descriptionContent += " and ";
                    }

                    if (segments[i].Contains("-")) {
                        var betweenSegmentDescription = this.GenerateBetweenSegmentDescription(segments[i], getRangeFormat, getSingleItemDescription);

                        //remove any leading comma
                        betweenSegmentDescription = betweenSegmentDescription.Replace(", ", "");

                        descriptionContent += betweenSegmentDescription;
                    }
                    else {
                        descriptionContent += getSingleItemDescription(segments[i]);
                    }
                }

                description = string.Format(getDescriptionFormat(expression), descriptionContent);
            }
            else if (expression.Contains("-")) {
                description = this.GenerateBetweenSegmentDescription(expression, getBetweenDescriptionFormat, getSingleItemDescription);
            }

            return description;
        }

        /// <summary>
        /// Generates the between segment description
        /// </summary>
        /// <param name="betweenExpression"></param>
        /// <param name="getBetweenDescriptionFormat"></param>
        /// <param name="getSingleItemDescription"></param>
        /// <returns>The between segment description</returns>
        protected string GenerateBetweenSegmentDescription(string betweenExpression, Func<string, string> getBetweenDescriptionFormat, Func<string, string> getSingleItemDescription) {
            var description = string.Empty;
            var betweenSegments = betweenExpression.Split('-');
            var betweenSegment1Description = getSingleItemDescription(betweenSegments[0]);
            var betweenSegment2Description = getSingleItemDescription(betweenSegments[1]);
            betweenSegment2Description = betweenSegment2Description.Replace(":00", ":59");
            var betweenDescriptionFormat = getBetweenDescriptionFormat(betweenExpression);
            description += string.Format(betweenDescriptionFormat, betweenSegment1Description, betweenSegment2Description);

            return description;
        }

        /// <summary>
        /// Given time parts, will construct a formatted time description
        /// </summary>
        /// <param name="hourExpression">Hours part</param>
        /// <param name="minuteExpression">Minutes part</param>
        /// <returns>Formatted time description</returns>
        protected string FormatTime(string hourExpression, string minuteExpression) {
            return this.FormatTime(hourExpression, minuteExpression, string.Empty);
        }

        /// <summary>
        /// Given time parts, will construct a formatted time description
        /// </summary>
        /// <param name="hourExpression">Hours part</param>
        /// <param name="minuteExpression">Minutes part</param>
        /// <param name="secondExpression">Seconds part</param>
        /// <returns>Formatted time description</returns>
        protected string FormatTime(string hourExpression, string minuteExpression, string secondExpression) {
            var hour = Convert.ToInt32(hourExpression);

            var period = string.Empty;
            if (!this.m_use24HourTimeFormat) {
                period = (hour >= 12) ? "PM" : "AM";
                if (period.Length > 0) {
                    // add preceding space
                    period = string.Concat(" ", period);
                }

                if (hour > 12) {
                    hour -= 12;
                }
                if (hour == 0) {
                    hour = 12;
                }
            }

            var minute = Convert.ToInt32(minuteExpression).ToString();
            var second = string.Empty;
            if (!string.IsNullOrEmpty(secondExpression)) {
                second = string.Concat(":", Convert.ToInt32(secondExpression).ToString().PadLeft(2, '0'));
            }

            return string.Format("{0}:{1}{2}{3}",
                hour.ToString().PadLeft(2, '0'), minute.PadLeft(2, '0'), second, period);
        }

        /// <summary>
        /// Transforms the verbosity of the expression description by stripping verbosity from original description
        /// </summary>
        /// <param name="description">The description to transform</param>
        /// <param name="isVerbose">If true, will leave description as it, if false, will strip verbose parts</param>
        /// <returns>The transformed description with proper verbosity</returns>
        protected string TransformVerbosity(string description, bool useVerboseFormat) {
            if (!useVerboseFormat) {
                description = description.Replace(", every minute", string.Empty);
                description = description.Replace(", every hour", string.Empty);
                description = description.Replace(", every day", string.Empty);
                description = Regex.Replace(description, @"\, ?$", "");
            }

            return description;
        }

        #region Static
        /// <summary>
        /// Generates a human readable string for the Cron Expression
        /// </summary>
        /// <param name="expression">The cron expression string</param>
        /// <returns>The cron expression description</returns>
        public static string? GetDescription(string expression) {
            return GetDescription(expression, new Options());
        }

        /// <summary>
        /// Generates a human readable string for the Cron Expression
        /// </summary>
        /// <param name="expression">The cron expression string</param>
        /// <param name="options">Options to control the output description</param>
        /// <returns>The cron expression description</returns>
        public static string? GetDescription(string expression, Options options) {
            var descriptor = new CronExpressionDescriptor(expression, options);
            return descriptor.GetDescription(DescriptionTypeEnum.FULL);
        }
        #endregion
    }

    /// <summary>
    /// Options for parsing and describing a Cron Expression
    /// </summary>
    internal class Options {
        internal Options() {
            this.ThrowExceptionOnParseError = true;
            this.Verbose = false;
            this.DayOfWeekStartIndexZero = true;
        }

        public bool ThrowExceptionOnParseError { get; set; }
        public bool Verbose { get; set; }
        public bool DayOfWeekStartIndexZero { get; set; }
        public bool? Use24HourTimeFormat { get; set; }
        public string Locale { get; set; } = "en-US";
    }
    internal enum DescriptionTypeEnum {
        FULL,
        TIMEOFDAY,
        SECONDS,
        MINUTES,
        HOURS,
        DAYOFWEEK,
        MONTH,
        DAYOFMONTH,
        YEAR
    }

    /// <summary>
    /// Cron Expression Parser
    /// </summary>
    internal class ExpressionParser {
        /* Cron reference

          ┌───────────── minute (0 - 59)
          │ ┌───────────── hour (0 - 23)
          │ │ ┌───────────── day of month (1 - 31)
          │ │ │ ┌───────────── month (1 - 12)
          │ │ │ │ ┌───────────── day of week (0 - 6) (Sunday to Saturday; 7 is also Sunday on some systems)
          │ │ │ │ │
          │ │ │ │ │
          │ │ │ │ │
          * * * * *  command to execute

         */

        private string m_expression;
        private Options m_options;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExpressionParser"/> class
        /// </summary>
        /// <param name="expression">The cron expression string</param>
        /// <param name="options">Parsing options</param>
        internal ExpressionParser(string expression, Options options) {
            this.m_expression = expression;
            this.m_options = options;
        }

        /// <summary>
        /// Parses the cron expression string
        /// </summary>
        /// <returns>A 7 part string array, one part for each component of the cron expression (seconds, minutes, etc.)</returns>
        public string[] Parse() {
            // Initialize all elements of parsed array to empty strings
            var parsed = new string[7].Select(el => "").ToArray();

            if (string.IsNullOrEmpty(this.m_expression)) {
#if NET_STANDARD_1X
        throw new Exception("Field 'expression' not found.");
#else
                throw new MissingFieldException("Field 'expression' not found.");
#endif
            }
            else {
                var expressionPartsTemp = this.m_expression.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                if (expressionPartsTemp.Length < 5) {
                    throw new FormatException(string.Format("Error: Expression only has {0} parts.  At least 5 part are required.", expressionPartsTemp.Length));
                }
                else if (expressionPartsTemp.Length == 5) {
                    //5 part cron so shift array past seconds element
                    Array.Copy(expressionPartsTemp, 0, parsed, 1, 5);
                }
                else if (expressionPartsTemp.Length == 6) {
                    /* We will detect if this 6 part expression has a year specified and if so we will shift the parts and treat the 
                       first part as a minute part rather than a second part. 

                       Ways we detect:
                         1. Last part is a literal year (i.e. 2020)
                         2. 3rd or 5th part is specified as "?" (DOM or DOW)
                    */
                    var isYearWithNoSecondsPart = Regex.IsMatch(expressionPartsTemp[5], "\\d{4}$") || expressionPartsTemp[4] == "?" || expressionPartsTemp[2] == "?";

                    if (isYearWithNoSecondsPart) {
                        // Shift parts over by one
                        Array.Copy(expressionPartsTemp, 0, parsed, 1, 6);
                    }
                    else {
                        Array.Copy(expressionPartsTemp, 0, parsed, 0, 6);
                    }
                }
                else if (expressionPartsTemp.Length == 7) {
                    parsed = expressionPartsTemp;
                }
                else {
                    throw new FormatException(string.Format("Error: Expression has too many parts ({0}).  Expression must not have more than 7 parts.", expressionPartsTemp.Length));
                }
            }

            this.NormalizeExpression(parsed);

            return parsed;
        }

        /// <summary>
        /// Converts cron expression components into consistent, predictable formats.
        /// </summary>
        /// <param name="expressionParts">A 7 part string array, one part for each component of the cron expression</param>
        private void NormalizeExpression(string[] expressionParts) {
            // Convert ? to * only for DOM and DOW
            expressionParts[3] = expressionParts[3].Replace("?", "*");
            expressionParts[5] = expressionParts[5].Replace("?", "*");

            // Convert 0/, 1/ to */
            if (expressionParts[0].StartsWith("0/")) {
                // Seconds
                expressionParts[0] = expressionParts[0].Replace("0/", "*/");
            }

            if (expressionParts[1].StartsWith("0/")) {
                // Minutes
                expressionParts[1] = expressionParts[1].Replace("0/", "*/");
            }

            if (expressionParts[2].StartsWith("0/")) {
                // Hours
                expressionParts[2] = expressionParts[2].Replace("0/", "*/");
            }

            if (expressionParts[3].StartsWith("1/")) {
                // DOM
                expressionParts[3] = expressionParts[3].Replace("1/", "*/");
            }

            if (expressionParts[4].StartsWith("1/")) {
                // Month
                expressionParts[4] = expressionParts[4].Replace("1/", "*/");
            }

            if (expressionParts[5].StartsWith("1/")) {
                // DOW
                expressionParts[5] = expressionParts[5].Replace("1/", "*/");
            }

            if (expressionParts[6].StartsWith("1/")) {
                // Years
                expressionParts[6] = expressionParts[6].Replace("1/", "*/");
            }

            // Adjust DOW based on dayOfWeekStartIndexZero option
            expressionParts[5] = Regex.Replace(expressionParts[5], @"(^\d)|([^#/\s]\d)", t => { //skip anything preceding by # or /
                var dowDigits = Regex.Replace(t.Value, @"\D", ""); // extract digit part (i.e. if "-2" or ",2", just take 2)
                var dowDigitsAdjusted = dowDigits;

                if (this.m_options.DayOfWeekStartIndexZero) {
                    // "7" also means Sunday so we will convert to "0" to normalize it
                    if (dowDigits == "7") {
                        dowDigitsAdjusted = "0";
                    }
                }
                else {
                    // If dayOfWeekStartIndexZero==false, Sunday is specified as 1 and Saturday is specified as 7.
                    // To normalize, we will shift the  DOW number down so that 1 becomes 0, 2 becomes 1, and so on.
                    dowDigitsAdjusted = (Int32.Parse(dowDigits) - 1).ToString();
                }

                return t.Value.Replace(dowDigits, dowDigitsAdjusted);
            });

            // Convert DOM '?' to '*'
            if (expressionParts[3] == "?") {
                expressionParts[3] = "*";
            }

            // Convert SUN-SAT format to 0-6 format
            for (var i = 0; i <= 6; i++) {
                var currentDay = (DayOfWeek)i;
                var currentDayOfWeekDescription = currentDay.ToString().Substring(0, 3).ToUpperInvariant();
                expressionParts[5] = Regex.Replace(expressionParts[5], currentDayOfWeekDescription, i.ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            }

            // Convert JAN-DEC format to 1-12 format
            for (var i = 1; i <= 12; i++) {
                var currentMonth = new DateTime(DateTime.Now.Year, i, 1);
                var currentMonthDescription = currentMonth.ToString("MMM", CultureInfo.InvariantCulture).ToUpperInvariant();
                expressionParts[4] = Regex.Replace(expressionParts[4], currentMonthDescription, i.ToString(), RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            }

            // Convert 0 second to (empty)
            if (expressionParts[0] == "0") {
                expressionParts[0] = string.Empty;
            }

            // If time interval is specified for seconds or minutes and next time part is single item, make it a "self-range" so
            // the expression can be interpreted as an interval 'between' range.
            //     For example:
            //     0-20/3 9 * * * => 0-20/3 9-9 * * * (9 => 9-9)
            //     */5 3 * * * => */5 3-3 * * * (3 => 3-3)
            if (expressionParts[2].IndexOfAny(new char[] { '*', '-', ',', '/' }) == -1
              && (Regex.IsMatch(expressionParts[1], @"\*|\/") || Regex.IsMatch(expressionParts[0], @"\*|\/"))) {
                expressionParts[2] += $"-{expressionParts[2]}";
            }

            // Loop through all parts and apply global normalization
            for (var i = 0; i < expressionParts.Length; i++) {
                // convert all '*/1' to '*'
                if (expressionParts[i] == "*/1") {
                    expressionParts[i] = "*";
                }

                /* Convert Month,DOW,Year step values with a starting value (i.e. not '*') to between expressions.
                   This allows us to reuse the between expression handling for step values.

                   For Example:
                    - month part '3/2' will be converted to '3-12/2' (every 2 months between March and December)
                    - DOW part '3/2' will be converted to '3-6/2' (every 2 days between Tuesday and Saturday)
                */

                if (expressionParts[i].Contains("/")
                    && expressionParts[i].IndexOfAny(new char[] { '*', '-', ',' }) == -1) {
                    string? stepRangeThrough = null;
                    switch (i) {
                        case 4: stepRangeThrough = "12"; break;
                        case 5: stepRangeThrough = "6"; break;
                        case 6: stepRangeThrough = "9999"; break;
                        default: stepRangeThrough = null; break;
                    }

                    if (stepRangeThrough != null) {
                        var parts = expressionParts[i].Split('/');
                        expressionParts[i] = string.Format("{0}-{1}/{2}", parts[0], stepRangeThrough, parts[1]);
                    }
                }
            }
        }
    }
}
