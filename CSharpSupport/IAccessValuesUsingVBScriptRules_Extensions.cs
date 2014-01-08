﻿using System;
using System.Collections.Generic;

namespace CSharpSupport
{
    public static class IAccessValuesUsingVBScriptRules_Extensions
	{
        public const int MaxNumberOfMemberAccessorBeforeArraysRequired = 5;

        public static object CALL(this IProvideVBScriptCompatFunctionality source, object target)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new string[0], new object[0]);
        }

        public static object CALL(this IProvideVBScriptCompatFunctionality source, object target, object[] arguments)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new string[0], arguments);
        }

        public static object CALL(this IProvideVBScriptCompatFunctionality source, object target, string member1, object[] arguments)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1 }, arguments);
        }

        public static object CALL(this IProvideVBScriptCompatFunctionality source, object target, string member1, string member2, object[] arguments)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1, member2 }, arguments);
        }

        public static object CALL(this IProvideVBScriptCompatFunctionality source, object target, string member1, string member2, string member3, object[] arguments)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1, member2, member3 }, arguments);
        }

        public static object CALL(this IProvideVBScriptCompatFunctionality source, object target, string member1, string member2, string member3, string member4, object[] arguments)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1, member2, member3, member4 }, arguments);
        }

        public static object CALL(this IProvideVBScriptCompatFunctionality source, object target, string member1, string member2, string member3, string member4, string member5, object[] arguments)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(target, new[] { member1, member2, member3, member4, member5 }, arguments);
        }
    }
}
