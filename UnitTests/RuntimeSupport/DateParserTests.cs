﻿using System;
using System.Collections.Generic;
using System.Globalization;
using VBScriptTranslator.RuntimeSupport;
using VBScriptTranslator.UnitTests.Shared;
using Xunit;

namespace VBScriptTranslator.UnitTests.RuntimeSupport
{
    public class DateParserTests
    {
        public class en_GB : CultureOverridingTests
        {
            public en_GB() : base(new CultureInfo("en-GB")) { }

            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, string value, int defaultYear, DateTime expectedDate)
            {
                Assert.Equal(expectedDate, (new DateParser(DateParser.DefaultMonthNameTranslator, defaultYearOverride: defaultYear)).Parse(value));
            }

            [Theory, MemberData("ErrorData")]
            public void ErrorCases(string description, string value, int defaultYear)
            {
                Assert.Throws<ArgumentException>(() =>
                {
                    (new DateParser(DateParser.DefaultMonthNameTranslator, defaultYearOverride: defaultYear)).Parse(value);
                });
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Day and month \"4 1\" (default year 2015)", "4 1", 2015, new DateTime(2015, 1, 4) };
                    yield return new object[] { "Day and month \"1 4\" (default year 2015)", "1 4", 2015, new DateTime(2015, 4, 1) };
                    yield return new object[] { "Day and month \"18 1\" (default year 2015)", "18 1", 2015, new DateTime(2015, 1, 18) };
                    yield return new object[] { "Day and month \"30 1\" (default year 2015)", "30 1", 2015, new DateTime(2015, 1, 30) };
                    yield return new object[] { "Day and month \"30 12\" (default year 2015)", "30 12", 2015, new DateTime(2015, 12, 30) };
                    yield return new object[] { "Month and day \"2 28\" (default year 2015)", "2 28", 2015, new DateTime(2015, 2, 28) };
                    yield return new object[] { "Month and year (no 29th Feb in 2015 so 29 must be the year) \"2 29\" (default year 2015)", "2 29", 2015, new DateTime(2029, 2, 1) };
                    yield return new object[] { "Month and day \"2 29\" (default year 2016)", "2 29", 2016, new DateTime(2016, 2, 29) };
                    yield return new object[] { "Month and year \"2 2015\"", "2 2015", 2016, new DateTime(2015, 2, 1) };
                    yield return new object[] { "Year and month (zero is only valid as a year, will be treated as 2000) \"0 1\"", "0 1", 2015, new DateTime(2000, 1, 1) };
                    yield return new object[] { "Month and year (zero is only valid as a year, will be treated as 2000) \"1 0\"", "1 0", 2015, new DateTime(2000, 1, 1) };

                    yield return new object[] { "Day, Month and Year \"1 2 9\"", "1 2 9", 2015, new DateTime(2009, 2, 1) };
                    yield return new object[] { "Day, Month and Year \"6 7 9\"", "6 7 9", 2015, new DateTime(2009, 7, 6) };
                    yield return new object[] { "Day, Month and Year \"17 2 9\"", "17 2 9", 2015, new DateTime(2009, 2, 17) };
                    yield return new object[] { "Month, Day and Year \"2 17 9\"", "2 17 9", 2015, new DateTime(2009, 2, 17) };
                    yield return new object[] { "Year, Month and Day \"2009-7-6\"", "2009-7-6", 2015, new DateTime(2009, 7, 6) };
                    yield return new object[] { "Year, Month and Day (zero is only valid as a year, will be treated as 2000) \"0 1 1\"", "0 1 1", 2015, new DateTime(2000, 1, 1) };
                    yield return new object[] { "Day, Month and Year (30 is interpreted 1930) \"5 10 30\"", "5 10 30", 2015, new DateTime(1930, 10, 5) };
                    yield return new object[] { "Day, Month and Year (29 is interpreted 2029) \"5 10 29\"", "5 10 29", 2015, new DateTime(2029, 10, 5) };

                    yield return new object[] { "Day and month name \"11 May\" (default year 2015)", "11 May", 2015, new DateTime(2015, 5, 11) };
                    yield return new object[] { "Month name and day \"May 11\" (default year 2015)", "May 11", 2015, new DateTime(2015, 5, 11) };
                    yield return new object[] { "Month name and short year \"May 0\"", "May 0", 2015, new DateTime(2000, 5, 1) };
                    yield return new object[] { "Month name and short year \"May 101\"", "May 101", 2015, new DateTime(101, 5, 1) };
                    yield return new object[] { "Month name and year \"May 2015\"", "May 2015", 2015, new DateTime(2015, 5, 1) };
                    yield return new object[] { "Short year and month name \"0 May\"", "0 May", 2015, new DateTime(2000, 5, 1) };
                    yield return new object[] { "Short year and month name \"101 May\"", "101 May", 2015, new DateTime(101, 5, 1) };
                    yield return new object[] { "Year and month name \"2015 May\"", "2015 May", 2015, new DateTime(2015, 5, 1) };

                    yield return new object[] { "Month name, day and year \"May 11 2015\"", "May 11 2015", 2015, new DateTime(2015, 5, 11) };
                    yield return new object[] { "Year, month name and day \"2015 May 11\"", "2015 May 11", 2015, new DateTime(2015, 5, 11) };
                    yield return new object[] { "Year, day and month name \"2015 11 May\"", "2015 11 May", 2015, new DateTime(2015, 5, 11) };
                    yield return new object[] { "Day, year and month name \"11 2015 May\"", "11 2015 May", 2015, new DateTime(2015, 5, 11) };
                    yield return new object[] { "Month name, year and day \"May 2015 11\"", "May 2015 11", 2015, new DateTime(2015, 5, 11) };

                    yield return new object[] { "Month name, day and short year \"May 11 15\"", "May 11 15", 2015, new DateTime(2015, 5, 11) };
                    yield return new object[] { "Day, month name and short year \"11 May 15\"", "11 May 15", 2015, new DateTime(2015, 5, 11) };
                    yield return new object[] { "Day, short year and month name \"11 15 May\"", "11 15 May", 2015, new DateTime(2015, 5, 11) };

                    yield return new object[] { "Full month name, day and year \"October 11 2015\"", "October 11 2015", 2015, new DateTime(2015, 10, 11) };
                    yield return new object[] { "Abbreviated month name, day and year \"Oct 11 2015\"", "Oct 11 2015", 2015, new DateTime(2015, 10, 11) };
                    yield return new object[] { "Lower-case abbreviated month name, day and year \"oct 11 2015\"", "oct 11 2015", 2015, new DateTime(2015, 10, 11) };
                    yield return new object[] { "Upper-case abbreviated month name, day and year \"OCT 11 2015\"", "OCT 11 2015", 2015, new DateTime(2015, 10, 11) };
                }
            }

            public static IEnumerable<object[]> ErrorData
            {
                get
                {
                    yield return new object[] { "Invalid month in yyyy-mm-dd \"2000-17-6\"", "2000-17-6", 2015 };
                    yield return new object[] { "Invalid day m-d-y \"2-29-6\"", "2-29-6", 2015 };
                    yield return new object[] { "Invalid double zero (zero is only valid as a year and a date can not consist of two years) \"0 0\"", "0 0", 2015 };
                    yield return new object[] { "Invalid month, year (zero is only valid as a year, if there are only two segments then the other must be a month) \"13 0\"", "13 0", 2015 };
                    yield return new object[] { "Invalid year, month (zero is only valid as a year, if there are only two segments then the other must be a month) \"0 13\"", "0 13", 2015 };
                    yield return new object[] { "Invalid day, month, year (zero is only valid as a year) \"0 0 2015\"", "0 0 2015", 2015 };
                    yield return new object[] { "Invalid day, month, year (zero is only valid as a year) \"0 1 2015\"", "0 1 2015", 2015 };
                }
            }
        }

        public class en_US : CultureOverridingTests
        {
            public en_US() : base(new CultureInfo("en-US")) { }

            [Theory, MemberData("SuccessData")]
            public void SuccessCases(string description, string value, int defaultYear, DateTime expectedDate)
            {
                Assert.Equal(expectedDate, (new DateParser(DateParser.DefaultMonthNameTranslator, defaultYearOverride: 2015)).Parse(value));
            }

            public static IEnumerable<object[]> SuccessData
            {
                get
                {
                    yield return new object[] { "Month and day \"4 1\" (default year 2015)", "4 1", 2015, new DateTime(2015, 4, 1) };
                    yield return new object[] { "Day and month (18 can't be a month since it's greater than 12) \"18 1\" (default year 2015)", "18 1", 2015, new DateTime(2015, 1, 18) };

                    yield return new object[] { "Month, Day and Year \"1 2 9\"", "1 2 9", 2015, new DateTime(2009, 1, 2) };
                    yield return new object[] { "Month, Day and Year \"6 7 9\"", "6 7 9", 2015, new DateTime(2009, 6, 7) };
                    yield return new object[] { "Year, Month and Day \"2009-7-6\"", "2009-7-6", 2015, new DateTime(2009, 7, 6) };
                    yield return new object[] { "Day, Month and Year (17 can't be a month since it's greater than 12) \"17 2 9\"", "17 2 9", 2015, new DateTime(2009, 2, 17) };
                    yield return new object[] { "Month, Day and Year \"2 17 9\"", "2 17 9", 2015, new DateTime(2009, 2, 17) };
                    yield return new object[] { "Month, Day and Year (30 is interpreted 1930) \"5 10 30\"", "5 10 30", 2015, new DateTime(1930, 5, 10) };
                    yield return new object[] { "Month, Day and Year (29 is interpreted 2029) \"5 10 29\"", "5 10 29", 2015, new DateTime(2029, 5, 10) };
                }
            }
        }
    }
}
