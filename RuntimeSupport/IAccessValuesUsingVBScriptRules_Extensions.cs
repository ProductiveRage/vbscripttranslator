using System;
using System.Collections.Generic;

namespace VBScriptTranslator.RuntimeSupport
{
    public static class IAccessValuesUsingVBScriptRules_Extensions
	{
        public const int MaxNumberOfMemberAccessorBeforeArraysRequired = 5;

        // Convenience methods so that the calling code can omit the "GetArgs" call if an IBuildCallArgumentProviders is already available (results in shorter
        // translated code)
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, IEnumerable<string> members, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException("argumentProviderBuilder");

            return source.CALL(context, target, members, argumentProviderBuilder.GetArgs());
        }
        public static void SET(this IAccessValuesUsingVBScriptRules source, object valueToSetTo, object context, object target, string optionalMemberAccessor, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException("argumentProviderBuilder");

            source.SET(valueToSetTo, context, target, optionalMemberAccessor, argumentProviderBuilder.GetArgs());
        }

        // This one allows for the arguments to not be mentioned at all if they're not required for a SET (unlike CALL, there is no concept of "forced brackets"
        // when there are zero arguments since a SET is always part of a value-setting statement, which means that the brackets are an essential part of the
        // statement and not optional tokens that may or may not be present)
        public static void SET(this IAccessValuesUsingVBScriptRules source, object valueToSetTo, object context, object target, string optionalMemberAccessor)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            source.SET(valueToSetTo, context, target, optionalMemberAccessor, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBrackets);
        }
        // This one allows for no arguments OR member accessors to be mentioned - this should only be used for errors cases (since, otherwise, a simple assignment
        // would be more appropriate, no SET call would be required at all). This may be used for the representation of "a = 1" where "a" is a function or a
        // constant, the translated output would be call to this function where the target to would actually be a call to RAISEERROR so that the valueToSet
        // may be evaluated and then a can-not-set-this error raised (consistent with how VBScript would handle it).
        public static void SET(this IAccessValuesUsingVBScriptRules source, object valueToSetTo, object context, object target)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            source.SET(valueToSetTo, context, target, null, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBrackets);
        }

        // Convenience methods for when there are no arguments (supporting up to MaxNumberOfMemberAccessorBeforeArraysRequired members accessors, just as the
        // extension methods further down do which look after the with-arguments signatures)
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(context, target, new string[0], ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBrackets);
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(context, target, new[] { member1 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBrackets);
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(context, target, new[] { member1, member2 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBrackets);
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(context, target, new[] { member1, member2, member3 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBrackets);
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, string member4)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(context, target, new[] { member1, member2, member3, member4 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBrackets);
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, string member4, string member5)
        {
            if (source == null)
                throw new ArgumentNullException("source");

            return source.CALL(context, target, new[] { member1, member2, member3, member4, member5 }, ZeroArgumentArgumentProvider.WithoutEnforcedArgumentBrackets);
        }

        // Convenience methods for when there are a known number of accessor members (including zero) and arguments - providing the argument builder means that
        // the translated code can be shorter (since there will be less "GetArgs" calls) but the trust is placed in these extension methods that the arguments
        // set will not be manipulated (extended). Since there would already trust that these won't manipulate any values if IProvideCallArguments references
        // were passed then this isn't a big deal (strictly speaking these methods express requirements greater than they really need but the shorter code is
        // worth it).
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException("argumentProviderBuilder");

            return source.CALL(context, target, new string[0], argumentProviderBuilder.GetArgs());
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException("argumentProviderBuilder");

            return source.CALL(context, target, new[] { member1 }, argumentProviderBuilder.GetArgs());
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException("argumentProviderBuilder");

            return source.CALL(context, target, new[] { member1, member2 }, argumentProviderBuilder.GetArgs());
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException("argumentProviderBuilder");

            return source.CALL(context, target, new[] { member1, member2, member3 }, argumentProviderBuilder.GetArgs());
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, string member4, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException("argumentProviderBuilder");

            return source.CALL(context, target, new[] { member1, member2, member3, member4 }, argumentProviderBuilder.GetArgs());
        }
        public static object CALL(this IAccessValuesUsingVBScriptRules source, object context, object target, string member1, string member2, string member3, string member4, string member5, IBuildCallArgumentProviders argumentProviderBuilder)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (argumentProviderBuilder == null)
                throw new ArgumentNullException("argumentProviderBuilder");

            return source.CALL(context, target, new[] { member1, member2, member3, member4, member5 }, argumentProviderBuilder.GetArgs());
        }

        private class ZeroArgumentArgumentProvider : IProvideCallArguments
        {
            public static IProvideCallArguments WithEnforcedArgumentBrackets = new ZeroArgumentArgumentProvider(true);
            public static IProvideCallArguments WithoutEnforcedArgumentBrackets = new ZeroArgumentArgumentProvider(false);
            private ZeroArgumentArgumentProvider(bool useBracketsWhereZeroArguments)
            {
                UseBracketsWhereZeroArguments = useBracketsWhereZeroArguments;
            }

            public int NumberOfArguments { get { return 0; } }

            public bool UseBracketsWhereZeroArguments { get; private set; }

            public IEnumerable<object> GetInitialValues()
            {
                return new object[0];
            }

            public void OverwriteValueIfByRef(int index, object value)
            {
                throw new ArgumentException("There are no arguments to overwrite");
            }
        }
    }
}
