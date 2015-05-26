using System;
using System.Text.RegularExpressions;

namespace CSharpSupport
{
    public class DateParser
    {
        private const string ONE_OR_MORE_DIGITS = @"\d+";
        private const string SOME_WHITESPACE_OR_NOTHING = @"\s*";
        private const string DATE_DELIMITER = @"[\/\-,\s]";
        private const string TIME_DELIMITER = @"[\.:]";

        private const string TWO_SEGMENT_DATE =
            "(" + ONE_OR_MORE_DIGITS + ")" +
            SOME_WHITESPACE_OR_NOTHING +
            DATE_DELIMITER +
            SOME_WHITESPACE_OR_NOTHING +
            "(" + ONE_OR_MORE_DIGITS + ")";
        private const string THREE_SEGMENT_DATE =
            TWO_SEGMENT_DATE +
            SOME_WHITESPACE_OR_NOTHING +
            DATE_DELIMITER +
            SOME_WHITESPACE_OR_NOTHING +
            "(" + ONE_OR_MORE_DIGITS + ")";
        private const string TWO_SEGMENT_TIME =
            "(" + ONE_OR_MORE_DIGITS + ")" +
            SOME_WHITESPACE_OR_NOTHING +
            TIME_DELIMITER +
            SOME_WHITESPACE_OR_NOTHING +
            "(" + ONE_OR_MORE_DIGITS + ")";
        private const string THREE_SEGMENT_TIME =
            TWO_SEGMENT_TIME +
            SOME_WHITESPACE_OR_NOTHING +
            TIME_DELIMITER +
            SOME_WHITESPACE_OR_NOTHING +
            "(" + ONE_OR_MORE_DIGITS + ")";

        private static readonly Regex _endOfStringThreeSegmentTimeComponent = new Regex(THREE_SEGMENT_TIME + "$", RegexOptions.Compiled);
        private static readonly Regex _endOfStringTwoSegmentTimeComponent = new Regex(TWO_SEGMENT_TIME + "$", RegexOptions.Compiled);
        private static readonly Regex _endOfStringSingleSegmentTimeComponent = new Regex("(" + ONE_OR_MORE_DIGITS + ")$", RegexOptions.Compiled);

        private static readonly Regex _wholeStringThreeSegmentDateComponent = new Regex("^" + THREE_SEGMENT_DATE + "$", RegexOptions.Compiled);
        private static readonly Regex _wholeStringTwoSegmentDateComponent = new Regex("^" + TWO_SEGMENT_DATE + "$", RegexOptions.Compiled);

        private static DateParser _default = new DateParser(() => DateTime.Now.Year);
        public static DateParser Default { get { return _default; } }

        /// <summary>
        /// This constructor should only be used for testing purposes, the static Default instance is appropriate for other uses
        /// </summary>
        public DateParser(int defaultYearOverride) : this(() => defaultYearOverride)
        {
            if ((defaultYearOverride < VBScriptConstants.EarliestPossibleDate.Year) || (defaultYearOverride > VBScriptConstants.LatestPossibleDate.Year))
                throw new ArgumentOutOfRangeException("defaultYearOverride must be a value that VBScript can represent");
        }
        private readonly Func<int> _defaultYearRetriever;
        private DateParser(Func<int> defaultYearRetriever)
        {
            if (defaultYearRetriever == null)
                throw new ArgumentNullException("defaultYearRetriever");
            
            _defaultYearRetriever = defaultYearRetriever;
        }

        /// <summary>
        /// This will throw an exception if the value can not be interpreted as a DateTime following VBScript's rules or if the value is null. Note that this ONLY supports
        /// the parsing of a string that is in a supported date format, it does not deal with cases such as CDate("2015"), where the string is parsed into a number and
        /// then a date calculated by taking the number of days from VBScript's "zero date". This will never return null;
        /// </summary>
        public DateResult Parse(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Null/blank value specified");

            TimeSpan time;
            value = ExtractAnyTimeComponent(value, out time).Trim();

            var threeSegmentDateComponentMatch = _wholeStringThreeSegmentDateComponent.Match(value);
            if (threeSegmentDateComponentMatch.Success)
            {
                return DateResult.FixedYear(
                    GetDate(
                        int.Parse(threeSegmentDateComponentMatch.Groups[1].Value),
                        int.Parse(threeSegmentDateComponentMatch.Groups[2].Value),
                        int.Parse(threeSegmentDateComponentMatch.Groups[3].Value)
                    )
                    .Add(time)
                );
            }

            var twoSegmentDateComponentMatch = _wholeStringTwoSegmentDateComponent.Match(value);
            if (!twoSegmentDateComponentMatch.Success)
                throw new ArgumentException("Invalid date format");

            return
                GetDate(
                    int.Parse(twoSegmentDateComponentMatch.Groups[1].Value),
                    int.Parse(twoSegmentDateComponentMatch.Groups[2].Value)
                )
                .Add(time);
        }

        private static string ExtractAnyTimeComponent(string value, out TimeSpan extractedTime)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Null/blank value specified");

            value = value.Trim();
            bool specifiesAM, specifiesPM;
            if (value.EndsWith("am", StringComparison.OrdinalIgnoreCase))
            {
                specifiesAM = true;
                specifiesPM = false;
                value = value.Substring(0, value.Length - 2).Trim();
            }
            else if (value.EndsWith("pm", StringComparison.OrdinalIgnoreCase))
            {
                specifiesAM = false;
                specifiesPM = true;
                value = value.Substring(0, value.Length - 2).Trim();
            }
            else
            {
                specifiesAM = false;
                specifiesPM = false;
            }

            int hour, minute, second;
            var threeSegmentTimeComponentMatch = _endOfStringThreeSegmentTimeComponent.Match(value);
            if (threeSegmentTimeComponentMatch.Success)
            {
                hour = int.Parse(threeSegmentTimeComponentMatch.Groups[1].Value);
                minute = int.Parse(threeSegmentTimeComponentMatch.Groups[2].Value);
                second = int.Parse(threeSegmentTimeComponentMatch.Groups[3].Value);
                value = value.Substring(0, value.Length - threeSegmentTimeComponentMatch.Groups[0].Value.Length).Trim();
            }
            else
            {
                var twoSegmentTimeComponentMatch = _endOfStringTwoSegmentTimeComponent.Match(value);
                if (twoSegmentTimeComponentMatch.Success)
                {
                    hour = int.Parse(twoSegmentTimeComponentMatch.Groups[1].Value);
                    minute = int.Parse(twoSegmentTimeComponentMatch.Groups[2].Value);
                    value = value.Substring(0, value.Length - twoSegmentTimeComponentMatch.Groups[0].Value.Length).Trim();
                }
                else
                {
                    if (specifiesAM || specifiesPM)
                    {
                        var singleSegmentTimeComponentMatch = _endOfStringSingleSegmentTimeComponent.Match(value);
                        if (!singleSegmentTimeComponentMatch.Success)
                        {
                            throw new ArgumentException(string.Format(
                                "Invalid date format, no time component could be extracted despite the presence of the {0} suffix",
                                specifiesAM ? "AM" : "PM"
                            ));
                        }
                        hour = int.Parse(singleSegmentTimeComponentMatch.Groups[1].Value);
                        value = value.Substring(0, value.Length - singleSegmentTimeComponentMatch.Groups[0].Value.Length).Trim();
                    }
                    else
                    {
                        hour = 0;
                    }
                    minute = 0;
                }
                second = 0;
            }
            if ((hour < 0) || (hour > 23) || (minute < 0) || (minute > 59) || (second < 0) || (second > 59))
            {
                throw new ArgumentException(string.Format(
                    "Invalid date format, time component out of range - indicates {0:00}:{1:00}:{2:00}",
                    hour,
                    minute,
                    second
                ));
            }
            if ((hour <= 12) && specifiesPM)
                hour += 12;
            extractedTime = new TimeSpan(hour, minute, second);
            return value;
        }

        /// <summary>
        /// This will throw an exception for an out-of-range month (1 to 12, inclusive) or year (must be within the range that VBScript can represent)
        /// </summary>
        private static int GetNumberOfDaysInMonth(int month, int year)
        {
            if ((month < 1) || (month > 12))
                throw new ArgumentOutOfRangeException("month");
            if ((year < VBScriptConstants.EarliestPossibleDate.Year) || (year > VBScriptConstants.LatestPossibleDate.Year))
                throw new ArgumentOutOfRangeException("year");

            return new DateTime(year, month, 1).AddMonths(1).AddDays(-1).Day;
        }

        /// <summary>
        /// This will throw an exception if the segments could not be parsed into a DateTime
        /// </summary>
        private static DateTime GetDate(int dateSegment1, int dateSegment2, int dateSegment3)
        {
            if (dateSegment1 < 0)
                throw new ArgumentOutOfRangeException("dateSegment1");
            if (dateSegment2 < 0)
                throw new ArgumentOutOfRangeException("dateSegment2");
            if (dateSegment3 < 0)
                throw new ArgumentOutOfRangeException("dateSegment3");

            // If the first two values could be days or months, then it's either d/m/y or m/d/y, depending upon current culture
            if ((dateSegment1 >= 1) && (dateSegment1 <= 12) && (dateSegment2 >= 1) && (dateSegment2 <= 12))
            {
                int day, month;
                if (PreferMonthBeforeDate())
                {
                    month = dateSegment1;
                    day = dateSegment2;
                }
                else
                {
                    day = dateSegment1;
                    month = dateSegment2;
                }
                var year = EnsureIsFourDigitYear(dateSegment3);
                if (day <= GetNumberOfDaysInMonth(month, year))
                    return new DateTime(year, month, day);
            }

            // If the first value may be a day but couldn't be a month and the second value could be a month, then it's d/m/y
            if ((dateSegment1 > 12) && (dateSegment1 <= 31) && (dateSegment2 >= 1) && (dateSegment2 <= 12))
            {
                var day = dateSegment1;
                var month = dateSegment2;
                var year = EnsureIsFourDigitYear(dateSegment3);
                if (day <= GetNumberOfDaysInMonth(month, year))
                    return new DateTime(year, month, day);
            }

            // If the second value may be a day but couldn't be a month and the first value could be a month, then it's m/d/y
            if ((dateSegment1 >= 1) && (dateSegment1 <= 12) && (dateSegment2 > 12) && (dateSegment2 <= 31))
            {
                var day = dateSegment2;
                var month = dateSegment1;
                var year = EnsureIsFourDigitYear(dateSegment3);
                if (day <= GetNumberOfDaysInMonth(month, year))
                    return new DateTime(year, month, day);
            }

            // So now we know that the first two segments can not be day / month or month / day, the only remaining valid configuration is y/m/d
            if ((dateSegment2 >= 1) && (dateSegment2 <= 12) && (dateSegment3 >= 1) && (dateSegment3 <= 31))
            {
                var year = EnsureIsFourDigitYear(dateSegment1);
                var month = dateSegment2;
                var day = dateSegment3;
                if (day <= GetNumberOfDaysInMonth(month, year))
                    return new DateTime(year, month, day);
            }
            throw new ArgumentException("Invalid date format");
        }

        /// <summary>
        /// If there are only two date segments (eg. "2 10" or "2 2015") then try to generate a value for the third segment, based upon the logic that VBScript would apply
        /// (for example "2 10" is 2nd October 2015 - if the current year is 2015 - and "2 2015" is 1st February 2015) and then return a DateTime with the now-three values.
        /// If the date segments provided are invalid then this will throw an exception.
        /// </summary>
        private DateResult GetDate(int dateSegment1, int dateSegment2)
        {
            if (dateSegment1 < 0)
                throw new ArgumentOutOfRangeException("dateSegment1");
            if (dateSegment2 < 0)
                throw new ArgumentOutOfRangeException("dateSegment2");

            // If there are only two segments then one must be the month and so at least one of the values must be less than 12 (the other segment may represent the year or
            // the day, it depends on what the two values are)
            var dateSegment1CouldBeMonth = (dateSegment1 >= 1) && (dateSegment1 <= 12);
            var dateSegment2CouldBeMonth = (dateSegment2 >= 1) && (dateSegment2 <= 12);
            if (!dateSegment1CouldBeMonth && !dateSegment2CouldBeMonth)
                throw new ArgumentException("Invalid date format (if there are only two segments then one must be the month and both values are outside the 1-12 range)");

            // If the segments are within the appropriate ranges that they could both represent either a day or a month, then the year will get a default value of the current
            // year. The complication comes from the fact that some culture prefer month to be specified first (American only?) and some prefer the date first. There can be
            // no clues to indicate one way or another in the value being parsed so we'll use the system culture (via the PreferMonthBeforeDate function).
            if (dateSegment1CouldBeMonth && dateSegment2CouldBeMonth)
            {
                var defaultYear = _defaultYearRetriever();
                if (PreferMonthBeforeDate())
                    return DateResult.DynamicYear(new DateTime(defaultYear, month: dateSegment1, day: dateSegment2));
                return DateResult.DynamicYear(new DateTime(defaultYear, month: dateSegment2, day: dateSegment1));
            }

            // If there is one segment within the month range (1 to 12, inclusive) and one within the day range but obviously not the month range (so 13 to 31, inclusive,
            // but depending upon that value of the month) then these values are combined and the year will be considered to be the default value that must be provided.
            var smallerDateSegment = Math.Min(dateSegment1, dateSegment2);
            var largerDateSegment = Math.Max(dateSegment1, dateSegment2);
            if ((smallerDateSegment >= 1) && (smallerDateSegment <= 12) && (largerDateSegment > 12))
            {
                var defaultYear = _defaultYearRetriever();
                if (largerDateSegment <= GetNumberOfDaysInMonth(smallerDateSegment, defaultYear))
                    return DateResult.DynamicYear(new DateTime(defaultYear, month: smallerDateSegment, day: largerDateSegment));
            }

            // Finally, if one segment is within the month range and other clearly not within the day range, then the other must be a year. Values of 100 or greater are
            // treated as simple year values while those less than 100 are treated as two-digit representations of four-digit values; anything less than 30 is presumed to
            // be in this century and 30 or greater is presumed to be a 1900. So zero, for example, is 2000, 1 is 2001, 10 is 2010, 29 is 2029 but 30 is 1930 and 96 is
            // 1996.
            int yearValue, monthValue;
            if (dateSegment1CouldBeMonth)
            {
                monthValue = dateSegment1;
                yearValue = dateSegment2;
            }
            else
            {
                monthValue = dateSegment2;
                yearValue = dateSegment1;
            }
            return DateResult.FixedYear(new DateTime(EnsureIsFourDigitYear(yearValue), monthValue, 1));
        }

        private static int EnsureIsFourDigitYear(int year)
        {
            if (year < 0)
                throw new ArgumentOutOfRangeException("year", "must be a positive value");

            if (year >= 100)
                return year;
            else if (year < 30)
                return year + 2000;
            else
                return year + 1900;
        }

        private static readonly Regex _simpleThreeSegmentNumberExtractor = new Regex(@"(\d+)\D+(\d+)\D+(\d+)", RegexOptions.Compiled);
        private static bool PreferMonthBeforeDate()
        {
            // Using the DateTime's ToShortDateString method should result in the date being rendered as three numeric values in a particular order (d m y) or (m d y).
            // We'll use this to determine whether the current system culture prefers month to come first or not. I suspect that this is not quite foolproof, but I'm
            // not sure how to improve it at this time (plus it will deal with American, who prefer month-first, and "everyone else", who prefer day first, so I think
            // that the bases are covered for the vast majority of use cases.. for mine at least! :)
            var sampleDate = new DateTime(2015, 5, 1); // The month and day values must be different in this sample date, obviously!
            var dateValuesMatchResult = _simpleThreeSegmentNumberExtractor.Match(sampleDate.ToShortDateString());
            return dateValuesMatchResult.Success && (int.Parse(dateValuesMatchResult.Groups[1].Value) == sampleDate.Month);
        }

        public class DateResult
        {
            public static DateResult DynamicYear(DateTime value) { return new DateResult(value, yearIsDynamic: true); }
            public static DateResult FixedYear(DateTime value) { return new DateResult(value, yearIsDynamic: false); }
            private DateResult(DateTime value, bool yearIsDynamic)
            {
                Value = value;
                YearIsDynamic = yearIsDynamic;
                RequiresLeapYear = (value.Month == 2) && (value.Day == 29);
            }
            public DateTime Value { get; private set; }
            public bool YearIsDynamic { get; private set; }
            public bool RequiresLeapYear { get; private set; }
            public DateResult Add(TimeSpan value)
            {
                return new DateResult(Value.Add(value), YearIsDynamic);
            }
        }
    }
}
