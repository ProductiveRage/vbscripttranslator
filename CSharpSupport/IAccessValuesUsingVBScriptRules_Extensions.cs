using System;
using System.Collections.Generic;

namespace CSharpSupport
{
    public static class IAccessValuesUsingVBScriptRules_Extensions
	{
        public const int MaxNumberOfMemberAccessorBeforeArraysRequired = 5;

        public static object CALL(this IAccessValuesUsingVBScriptRules source, object target)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new string[0], new ZeroArgumentArgumentProvider());
        }

        public static object CALL(this IAccessValuesUsingVBScriptRules source, object target, IProvideCallArguments argumentProvider)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new string[0], argumentProvider);
        }

        public static object CALL(this IAccessValuesUsingVBScriptRules source, object target, string member1, IProvideCallArguments argumentProvider)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1 }, argumentProvider);
        }

        public static object CALL(this IAccessValuesUsingVBScriptRules source, object target, string member1, string member2, IProvideCallArguments argumentProvider)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1, member2 }, argumentProvider);
        }

        public static object CALL(this IAccessValuesUsingVBScriptRules source, object target, string member1, string member2, string member3, IProvideCallArguments argumentProvider)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1, member2, member3 }, argumentProvider);
        }

        public static object CALL(this IAccessValuesUsingVBScriptRules source, object target, string member1, string member2, string member3, string member4, IProvideCallArguments argumentProvider)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1, member2, member3, member4 }, argumentProvider);
        }

        public static object CALL(this IAccessValuesUsingVBScriptRules source, object target, string member1, string member2, string member3, string member4, string member5, IProvideCallArguments argumentProvider)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1, member2, member3, member4, member5 }, argumentProvider);
        }
        private class ZeroArgumentArgumentProvider : IProvideCallArguments
        {
            public int NumberOfArguments { get { return 0; } }

            public IEnumerable<object> GetInitialValues()
            {
                throw new ArgumentException("There are no arguments to retrieve values for");
            }

            public void OverwriteValueIfByRef(int index, object value)
            {
                throw new ArgumentException("There are no arguments to overwrite");
            }
        }
    }
}
