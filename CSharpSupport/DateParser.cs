using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace VBScriptTranslator.RuntimeSupport
{
    public class DateParser
    {
        private const string ONE_OR_MORE_DIGITS = @"\d+";
        private const string SOME_WHITESPACE_OR_NOTHING = @"\s*";
        private const string DATE_DELIMITER = @"[\/\-,\s]";
        private const string TIME_DELIMITER = @"[\.:]";
        private const string POTENTIAL_MONTH_NAME = @"[^\d|\s|\/|\-|,|\.|:]+"; // One or more characters that are not delimiters, whitespace or digits

        // Dates consisting of two or three numeric segments
        private const string TWO_NUMBER_DATE = "(" + ONE_OR_MORE_DIGITS + ")" + SOME_WHITESPACE_OR_NOTHING + DATE_DELIMITER + SOME_WHITESPACE_OR_NOTHING + "(" + ONE_OR_MORE_DIGITS + ")";
        private const string THREE_NUMBER_DATE = TWO_NUMBER_DATE + SOME_WHITESPACE_OR_NOTHING + DATE_DELIMITER + SOME_WHITESPACE_OR_NOTHING + "(" + ONE_OR_MORE_DIGITS + ")";

        // Dates consisting of a month name and one numeric segment
        private const string MONTHNAME_THEN_NUMBER_DATE = "(" + POTENTIAL_MONTH_NAME + ")" + SOME_WHITESPACE_OR_NOTHING + DATE_DELIMITER + SOME_WHITESPACE_OR_NOTHING + "(" + ONE_OR_MORE_DIGITS + ")";
        private const string NUMBER_THEN_MONTHNAME_DATE = "(" + ONE_OR_MORE_DIGITS + ")" + SOME_WHITESPACE_OR_NOTHING + DATE_DELIMITER + SOME_WHITESPACE_OR_NOTHING + "(" + POTENTIAL_MONTH_NAME + ")";

        // Dates consisting of a month name and two numeric segments
        private const string MONTHNAME_THEN_TWO_NUMBERS_DATE = MONTHNAME_THEN_NUMBER_DATE + SOME_WHITESPACE_OR_NOTHING + DATE_DELIMITER + SOME_WHITESPACE_OR_NOTHING + "(" + ONE_OR_MORE_DIGITS + ")";
        private const string NUMBER_THEN_MONTHNAME_THEN_NUMBERS_DATE = "(" + ONE_OR_MORE_DIGITS + ")" + SOME_WHITESPACE_OR_NOTHING + DATE_DELIMITER + SOME_WHITESPACE_OR_NOTHING + MONTHNAME_THEN_NUMBER_DATE;
        private const string TWO_NUMBERS_THEN_MONTHNAME_DATE = "(" + ONE_OR_MORE_DIGITS + ")" + SOME_WHITESPACE_OR_NOTHING + DATE_DELIMITER + SOME_WHITESPACE_OR_NOTHING + NUMBER_THEN_MONTHNAME_DATE;

        // Time matches, for either two or three numeric segments (any AM/PM content is removed before using regular expressions)
        private const string TWO_SEGMENT_TIME = "(" + ONE_OR_MORE_DIGITS + ")" + SOME_WHITESPACE_OR_NOTHING + TIME_DELIMITER + SOME_WHITESPACE_OR_NOTHING + "(" + ONE_OR_MORE_DIGITS + ")";
        private const string THREE_SEGMENT_TIME = TWO_SEGMENT_TIME + SOME_WHITESPACE_OR_NOTHING + TIME_DELIMITER + SOME_WHITESPACE_OR_NOTHING + "(" + ONE_OR_MORE_DIGITS + ")";

        private static readonly Regex _endOfStringThreeSegmentTimeComponent = new Regex(THREE_SEGMENT_TIME + "$", RegexOptions.Compiled);
        private static readonly Regex _endOfStringTwoSegmentTimeComponent = new Regex(TWO_SEGMENT_TIME + "$", RegexOptions.Compiled);
        private static readonly Regex _endOfStringSingleSegmentTimeComponent = new Regex("(" + ONE_OR_MORE_DIGITS + ")$", RegexOptions.Compiled);

        private static readonly Regex _wholeStringThreeNumericSegmentDateComponent = new Regex("^" + THREE_NUMBER_DATE + "$", RegexOptions.Compiled);
        private static readonly Regex _wholeStringTwoNumericSegmentDateComponent = new Regex("^" + TWO_NUMBER_DATE + "$", RegexOptions.Compiled);
        private static readonly Regex _wholeStringMonthNameThenNumberDateComponent = new Regex("^" + MONTHNAME_THEN_NUMBER_DATE + "$", RegexOptions.Compiled);
        private static readonly Regex _wholeStringNumberThenMonthNameDateComponent = new Regex("^" + NUMBER_THEN_MONTHNAME_DATE + "$", RegexOptions.Compiled);
        private static readonly Regex _wholeStringMonthNameThenTwoNumbersDateComponent = new Regex("^" + MONTHNAME_THEN_TWO_NUMBERS_DATE + "$", RegexOptions.Compiled);
        private static readonly Regex _wholeStringNumberThenMonthNameThenNumberDateComponent = new Regex("^" + NUMBER_THEN_MONTHNAME_THEN_NUMBERS_DATE + "$", RegexOptions.Compiled);
        private static readonly Regex _wholeStringTwoNumbersThenMonthNameDateComponent = new Regex("^" + TWO_NUMBERS_THEN_MONTHNAME_DATE + "$", RegexOptions.Compiled);

        private static DateParser _default = new DateParser(DefaultMonthNameTranslator, () => DateTime.Now.Year);
        public static DateParser Default { get { return _default; } }

        /// <summary>
        /// This constructor should only be used for testing purposes, the static Default instance is appropriate for most other uses
        /// </summary>
        public DateParser(Func<string, int> monthNameTranslator, int defaultYearOverride) : this(monthNameTranslator, () => defaultYearOverride)
        {
            if ((defaultYearOverride < VBScriptConstants.EarliestPossibleDate.Year) || (defaultYearOverride > VBScriptConstants.LatestPossibleDate.Year))
                throw new ArgumentOutOfRangeException("defaultYearOverride must be a value that VBScript can represent");
        }
        private readonly Func<string, int> _monthNameTranslator;
        private readonly Func<int> _defaultYearRetriever;
        private DateParser(Func<string, int> monthNameTranslator, Func<int> defaultYearRetriever)
        {
            if (monthNameTranslator == null)
                throw new ArgumentNullException("monthNameTranslator");
            if (defaultYearRetriever == null)
                throw new ArgumentNullException("defaultYearRetriever");

            _monthNameTranslator = monthNameTranslator;
            _defaultYearRetriever = defaultYearRetriever;
        }

        /// <summary>
        /// This translates month names using the current culture - it supports full and abbreviated month names
        /// </summary>
        public static int DefaultMonthNameTranslator(string monthName)
        {
            if (string.IsNullOrWhiteSpace(monthName))
                throw new ArgumentException("Null/blank monthName specified");

            DateTime date;
            if (DateTime.TryParseExact(monthName, "MMM", CultureInfo.CurrentCulture, DateTimeStyles.None, out date)
            || DateTime.TryParseExact(monthName, "MMMM", CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                return date.Month;
            throw new ArgumentException("Invalid monthName specified (for current culture \"" + CultureInfo.CurrentCulture.DisplayName + "\"): \"" + monthName + "\"");
        }

        /// <summary>
        /// Translate a numeric value into a date, following VBScript's logic. This will throw an OverflowException for a number outside of the acceptable range.
        /// </summary>
        public DateTime Parse(double value)
        {
            // VBScript has some absolutely bonkers logic for negative values here - eg. CDate(-400.2) = 1898-11-25 04:48:00 which is equal to
            // (CDate(-400) + CDate(0.2)) and NOT equal to (CDate(-400) + CDate(-0.2)). It appears that the negative sign is just removed from
            // fractions, since CDate(0.1) = CDate(-0.1) and CDate(0.2) = CDate(-0.2) in VBScript.
            var integerPortion = Math.Truncate(value);
            var isGreatestPossibleDate = (integerPortion == Math.Truncate(VBScriptConstants.LatestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays));
            double fractionalPortion;
            if (isGreatestPossibleDate)
            {
                // There is also some even stranger logic around times on the last possible day that can be represented. A time component from
                // a .9 value (eg. -100.9, 0.9, 10.9, 42140.9) will always show 21:36:00 EXCEPT for when part of a value that represents that
                // time on the very last representable day, in which case it will 21:35:59 (try it: 2958465.9). VBScript does not seem to lose
                // any precision when converting from and to dates - eg. CDate(2958465.9) is "31/12/9999 21:35:59" and CDbl(CDate(2958465.9))
                // is still 2958465.9 - so I think that there must be a bug in the date-handling that I don't know how to fully recreate. I'm
                // going to try to always return a value from here that will be consistent with what VBScript would return from CDate, so the
                // value 2958465.9 will return "31/12/9999 21:35:59" (while all other .9 values will return 21:36:00) but at the sacrifice of
                // "back and forth precision" - so CDbl(CDate(2958465.9)) will be 2958465.8999884259, though all other values will maintain
                // precision correctly (so CDbl(CDate(2958464.9)) will be 2958464.9)
                // - On top of this, there is a hard limit on the time component of the last possible day, at which point an overflow will
                //   occur; 2958465.9999999997672 will overflow while 2958465.99999999976719999999 (as many 9s as you like) won't. I have
                //   no idea what that relates to, but a special case is applied here to deal with it.
                fractionalPortion = Math.Abs(value - integerPortion);
                if (fractionalPortion >= 0.9999999997672)
                    throw new OverflowException();
                var numberOfDigitsToAllow = 8;
                fractionalPortion = Math.Truncate(fractionalPortion * Math.Pow(10, numberOfDigitsToAllow)) / Math.Pow(10, numberOfDigitsToAllow);
            }
            else
                fractionalPortion = Math.Abs(value - integerPortion);
            var isEarliestPossibleDate = (integerPortion == Math.Truncate(VBScriptConstants.EarliestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays));
            if (isEarliestPossibleDate)
            {
                // There is a similar limit to the time component on the other end of the scale, at which point an overflow will occur (the
                // value -657434.0.9999999999418 will CDate while -657434.9999999999417999999  99, with as many 9s as you like, will not)
                if (fractionalPortion >= 0.9999999999418)
                    throw new OverflowException();
            }
            if ((integerPortion > Math.Truncate(VBScriptConstants.LatestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays))
            || (integerPortion < Math.Truncate(VBScriptConstants.EarliestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays)))
                throw new OverflowException();
            var calculatedTimeComponent = VBScriptConstants.ZeroDate.AddDays(fractionalPortion);
            if (isGreatestPossibleDate)
            {
                // Continuing the crazy-logic-on-last-representable-date, only must the precision of the time component be reduced on the
                // greatest possible date, but any millisecond component must be stripped (not rounded) completely. The is how we ensure
                // that 2958465.9 results in the time "21:35:59" and not "21:36:00".
                calculatedTimeComponent = calculatedTimeComponent.Subtract(TimeSpan.FromMilliseconds(calculatedTimeComponent.Millisecond));
            }
            return calculatedTimeComponent.AddDays(integerPortion);
        }

        /// <summary>
        /// This will throw an exception if the value can not be interpreted as a DateTime following VBScript's rules or if the value is null. Note that this ONLY supports
        /// the parsing of a string that is in a supported date format, it does not deal with cases such as CDate("2015"), where the string is parsed into a number and
        /// then a date calculated by taking the number of days from VBScript's "zero date". If a date outside of VBScript's expressible range is described then an
        /// OverflowException will be thrown. This will never return null.
        /// </summary>
        public DateTime Parse(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Null/blank value specified");

            TimeSpan time;
            value = ExtractAnyTimeComponent(value, out time).Trim();
            
            var date = ParseDateOnly(value);
            if ((date < VBScriptConstants.EarliestPossibleDate.Date) || (date > VBScriptConstants.LatestPossibleDate.Date))
                throw new OverflowException();

            return date.Add(time);
        }

        private DateTime ParseDateOnly(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Null/blank value specified");

            // First, check for numeric date formats - these can be validated fully here (and clearly invalid dates will result in an exception)
            var threeSegmentDateComponentMatch = _wholeStringThreeNumericSegmentDateComponent.Match(value);
            if (threeSegmentDateComponentMatch.Success)
            {
                return GetDate(
                    int.Parse(threeSegmentDateComponentMatch.Groups[1].Value),
                    int.Parse(threeSegmentDateComponentMatch.Groups[2].Value),
                    int.Parse(threeSegmentDateComponentMatch.Groups[3].Value)
                );
            }
            var twoSegmentDateComponentMatch = _wholeStringTwoNumericSegmentDateComponent.Match(value);
            if (twoSegmentDateComponentMatch.Success)
            {
                return GetDate(
                    int.Parse(twoSegmentDateComponentMatch.Groups[1].Value),
                    int.Parse(twoSegmentDateComponentMatch.Groups[2].Value)
                );
            }

            // Now we have to check for the with-month-name date formats
            // - If there is a month name and a single numeric value, then prefer the value to be a date (using the default year), VBScript does not take the order of the
            //   values to be of any significance. If the value does not fall within the acceptable range then assume it's a year (and default to the 1st of the month).
            string monthNameFromTwoSegmentFormat;
            int numericValueFromTwoSegmentFormat;
            var monthNameThenNumberComponentMatch = _wholeStringMonthNameThenNumberDateComponent.Match(value);
            if (monthNameThenNumberComponentMatch.Success)
            {
                monthNameFromTwoSegmentFormat = monthNameThenNumberComponentMatch.Groups[1].Value;
                numericValueFromTwoSegmentFormat = int.Parse(monthNameThenNumberComponentMatch.Groups[2].Value);
            }
            else
            {
                var numberThenMonthNameComponentMatch = _wholeStringNumberThenMonthNameDateComponent.Match(value);
                if (numberThenMonthNameComponentMatch.Success)
                {
                    monthNameFromTwoSegmentFormat = numberThenMonthNameComponentMatch.Groups[2].Value;
                    numericValueFromTwoSegmentFormat = int.Parse(numberThenMonthNameComponentMatch.Groups[1].Value);
                }
                else
                {
                    monthNameFromTwoSegmentFormat = null;
                    numericValueFromTwoSegmentFormat = 0;
                }
            }
            if (monthNameFromTwoSegmentFormat != null)
            {
                var month = _monthNameTranslator(monthNameFromTwoSegmentFormat);
                var defaultYear = _defaultYearRetriever();
                if ((numericValueFromTwoSegmentFormat >= 1) && (numericValueFromTwoSegmentFormat <= GetNumberOfDaysInMonth(month, defaultYear)))
                    return new DateTime(defaultYear, month, numericValueFromTwoSegmentFormat);
                return new DateTime(EnsureIsFourDigitYear(numericValueFromTwoSegmentFormat), month, 1);
            }
            // - If there is a month name and two numeric values and there is would be an ambiguity in what the numbers represent (the ordering of the segments is given
            //   no significance by VBScript) then it is assumed that the first value is the date and the second the year. If this combination is not valid then the
            //   first is the year and the second the date (if THIS is not valid then the date literal is invalid).
            string monthNameFromThreeSegmentFormat;
            int firstNumericValueFromThreeSegmentFormat, secondNumericValueFromThreeSegmentFormat;
            var monthNameThenTwoNumbersComponentMatch = _wholeStringMonthNameThenTwoNumbersDateComponent.Match(value);
            if (monthNameThenTwoNumbersComponentMatch.Success)
            {
                monthNameFromThreeSegmentFormat = monthNameThenTwoNumbersComponentMatch.Groups[1].Value;
                firstNumericValueFromThreeSegmentFormat = int.Parse(monthNameThenTwoNumbersComponentMatch.Groups[2].Value);
                secondNumericValueFromThreeSegmentFormat = int.Parse(monthNameThenTwoNumbersComponentMatch.Groups[3].Value);
            }
            else
            {
                var numberThenMonthNameThenNumberComponentMatch = _wholeStringNumberThenMonthNameThenNumberDateComponent.Match(value);
                if (numberThenMonthNameThenNumberComponentMatch.Success)
                {
                    monthNameFromThreeSegmentFormat = numberThenMonthNameThenNumberComponentMatch.Groups[2].Value;
                    firstNumericValueFromThreeSegmentFormat = int.Parse(numberThenMonthNameThenNumberComponentMatch.Groups[1].Value);
                    secondNumericValueFromThreeSegmentFormat = int.Parse(numberThenMonthNameThenNumberComponentMatch.Groups[3].Value);
                }
                else
                {
                    var twoNumbersThenMonthNameComponentMatch = _wholeStringTwoNumbersThenMonthNameDateComponent.Match(value);
                    if (twoNumbersThenMonthNameComponentMatch.Success)
                    {
                        monthNameFromThreeSegmentFormat = twoNumbersThenMonthNameComponentMatch.Groups[3].Value;
                        firstNumericValueFromThreeSegmentFormat = int.Parse(twoNumbersThenMonthNameComponentMatch.Groups[1].Value);
                        secondNumericValueFromThreeSegmentFormat = int.Parse(twoNumbersThenMonthNameComponentMatch.Groups[2].Value);
                    }
                    else
                    {
                        monthNameFromThreeSegmentFormat = null;
                        firstNumericValueFromThreeSegmentFormat = secondNumericValueFromThreeSegmentFormat = 0;
                    }
                }
            }
            if (monthNameFromThreeSegmentFormat != null)
            {
                var month = _monthNameTranslator(monthNameFromThreeSegmentFormat);
                var date = firstNumericValueFromThreeSegmentFormat;
                var year = EnsureIsFourDigitYear(secondNumericValueFromThreeSegmentFormat);
                if ((date >= 1) && (date <= GetNumberOfDaysInMonth(month, year)))
                    return new DateTime(year, month, date);
                date = secondNumericValueFromThreeSegmentFormat;
                year = EnsureIsFourDigitYear(firstNumericValueFromThreeSegmentFormat);
                if ((date >= 1) && (date <= GetNumberOfDaysInMonth(month, year)))
                    return new DateTime(year, month, date);
            }

            throw new ArgumentException("Invalid date format");
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
        private DateTime GetDate(int dateSegment1, int dateSegment2)
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
                    return new DateTime(defaultYear, month: dateSegment1, day: dateSegment2);
                return new DateTime(defaultYear, month: dateSegment2, day: dateSegment1);
            }

            // If there is one segment within the month range (1 to 12, inclusive) and one within the day range but obviously not the month range (so 13 to 31, inclusive,
            // but depending upon that value of the month) then these values are combined and the year will be considered to be the default value that must be provided.
            var smallerDateSegment = Math.Min(dateSegment1, dateSegment2);
            var largerDateSegment = Math.Max(dateSegment1, dateSegment2);
            if ((smallerDateSegment >= 1) && (smallerDateSegment <= 12) && (largerDateSegment > 12))
            {
                var defaultYear = _defaultYearRetriever();
                if (largerDateSegment <= GetNumberOfDaysInMonth(smallerDateSegment, defaultYear))
                    return new DateTime(defaultYear, month: smallerDateSegment, day: largerDateSegment);
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
            return new DateTime(EnsureIsFourDigitYear(yearValue), monthValue, 1);
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
    }
}
