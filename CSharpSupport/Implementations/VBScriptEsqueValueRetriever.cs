using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using CSharpSupport.Attributes;
using CSharpSupport.Exceptions;

namespace CSharpSupport.Implementations
{
    /// <summary>
    /// When this has to use reflection to inspect a target class to determine how to access a member function or property, it caches that lookup and so
    /// subsequent requests for the same type and member will be faster. The intention is that an instance of this be used across all requests that may
    /// be executing and so is written to be thread safe.
    /// </summary>
    public class VBScriptEsqueValueRetriever : IAccessValuesUsingVBScriptRules
    {
        // It's feasible that access to the invoker cacher will need to support multi-threaded access depending upon the application in question, so the
        // ConcurrentDictionary seems like a good choice. The most common cases I'm envisaging are for the cache to be built up as various execution paths
        // are followed and then for it to essentially become full. The ConcurrentDictionary allows lock-free reading and so doesn't introduce any costs
        // once this situation is reached.
        private readonly Func<string, string> _nameRewriter;
        private readonly ConcurrentDictionary<InvokerCacheKey, GetInvoker> _getInvokerCache;
        private readonly ConcurrentDictionary<InvokerCacheKey, SetInvoker> _setInvokerCache;
        public VBScriptEsqueValueRetriever(Func<string, string> nameRewriter)
        {
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            _nameRewriter = nameRewriter;
            _getInvokerCache = new ConcurrentDictionary<InvokerCacheKey, GetInvoker>();
            _setInvokerCache = new ConcurrentDictionary<InvokerCacheKey, SetInvoker>();
        }

        /// <summary>
        /// Reduce a reference down to a value type, applying VBScript defaults logic - thrown an exception if this is not possible (null is
        /// acceptable as an input and corresponding return value)
        /// </summary>
        public object VAL(object o, string optionalExceptionMessageForInvalidContent = null)
        {
            if (IsVBScriptValueType(o))
                return o;

            if (IsVBScriptNothing(o))
                throw new ObjectVariableNotSetException(optionalExceptionMessageForInvalidContent);

            var defaultValueFromObject = InvokeGetter(o, null, new object[0], onlyConsiderMethods: false);
            if (IsVBScriptValueType(defaultValueFromObject))
                return defaultValueFromObject;

            // We don't recursively try defaults, so if this default is still not a value type then we're out of luck
            throw new ObjectVariableNotSetException(optionalExceptionMessageForInvalidContent + " Object expected (default method/property of object also returned non-value type data)");
        }

        /// <summary>
        /// This will never throw an exception, a value is either considered by VBScript to be a value type (including values such as Empty,
        /// Null, numbers, dates, strings, arrays) or not
        /// </summary>
        public bool IsVBScriptValueType(object o)
        {
            return ((o == null) || (o == DBNull.Value) || (o is ValueType) || (o is string) || (o.GetType().IsArray));
        }

        /// <summary>
        /// The comparison (o == VBScriptConstants.Nothing) will return false even if o is VBScriptConstants.Nothing due to the implementation details of
        /// DispatchWrapper. This method delivers a reliable way to test for it.
        /// </summary>
        private bool IsVBScriptNothing(object o)
        {
            return ((o is DispatchWrapper) && ((DispatchWrapper)o).WrappedObject == null);
        }

        /// <summary>
        /// This will only return a non-VBScript-value-type, if unable to then an exception will be raised (this is used to wrap the right-hand
        /// side of a SET assignment)
        /// </summary>
        public object OBJ(object o, string optionalExceptionMessageForInvalidContent = null)
        {
            if ((o == null) || IsVBScriptValueType(o))
                throw new ObjectRequiredException(optionalExceptionMessageForInvalidContent);

            return o;
        }

        /// <summary>
        /// Reduce a reference down to a boolean value type, in the same ways that VBScript would attempt.
        /// </summary>
        public bool BOOL(object o, string optionalExceptionMessageForInvalidContent = null)
        {
            o = VAL(o, optionalExceptionMessageForInvalidContent);
            if (o == null)
                return false;
            if (o == DBNull.Value)
                throw new InvalidUseOfNullException(optionalExceptionMessageForInvalidContent);
            if (o is bool)
                return (bool)o;
            if (o is DateTime)
                return ((DateTime)o) != VBScriptConstants.ZeroDate;
            var valueString = o.ToString();
            if (valueString.Equals("true", StringComparison.OrdinalIgnoreCase))
                return true;
            if (valueString.Equals("false", StringComparison.OrdinalIgnoreCase))
                return false;
            double valueNumber;
            if (!double.TryParse(valueString, out valueNumber))
                throw new TypeMismatchException(optionalExceptionMessageForInvalidContent);
            return valueNumber != 0;
        }

        /// <summary>
        /// This wraps a call to NUM and allows an exception to be made for DBNull.Value (VBScript Null) in that the same value will be returned
        /// (it is not a valid input for NUM).
        /// </summary>
        public object NullableNUM(object o)
        {
            o = VAL(o);
            return (o == DBNull.Value) ? DBNull.Value : NUM(o);
        }

        /// <summary>
        /// Reduce a reference down to a numeric value type, applying VBScript defaults logic and then trying to parse as a number - throwing
        /// an exception if this is not possible. Null (aka VBScript Empty) is acceptable and will result in zero being returned. DBNull.Value
        /// (aka VBScript Null) is not acceptable and will result in an exception being raised, as any other invalid value (eg. a string or
        /// an object without an appropriate default property member) will. This is used by translated code and is similar in many ways to the
        /// IProvideVBScriptCompatFunctionalityToIndividualRequests.CDBL method, but it will not always return a double - it may return Int16,
        /// Int32, DateTime or other types. If there are numericValuesTheTypeMustBeAbleToContain values specified, then each of these will be
        /// passed through NUM as well and then the returned value's type will be such that it can contain all of those values (eg. if o is
        /// 1 and there are no numericValuesTheTypeMustBeAbleToContain then an Int16 will be returned, but if the return value must also be
        /// able to contain 32,768 then an Int32 representation will be returned. This means that this function may throw an overflow
        /// exception - if, for example, o is a date and it is asked to contain a numeric value that is would result in a date outside of
        /// the VBScript supported range then an overflow exception would be raised).
        /// </summary>
        public object NUM(object o, params object[] numericValuesTheTypeMustBeAbleToContain)
        {
            // Before we try anything, we need to call VAL to ensure we don't have something VBScript wouldn't consider a "value type"
            o = VAL(o);

            // Now, check whether it's a DateTime. If so, then don't try any harder to extract a value from it - VBScript considers DateTime
            // to be a numeric type.
            object valueToConvert;
            if (o is DateTime)
                valueToConvert = o;
            else
            {
                // Next, try to force it into one of the VBScript-acceptable number types - if it fails then it's a type mismatch
                try
                {
                    valueToConvert = GetAsVBScriptNumber(o);
                }
                catch (SpecificVBScriptException)
                {
                    throw;
                }
                catch (Exception e)
                {
                    throw new TypeMismatchException(e);
                }
            }

            // Now we have a numeric value, we need to check numericValuesTheTypeMustBeAbleToContain. They all must be parseable as numbers.
            // If there are no related values or if all of the related value have the same type as the primary value then nothing has to
            // change. Otherwise we need to ensure that the type we return can contain them all..
            var relatedNumericValues = (numericValuesTheTypeMustBeAbleToContain ?? new object[0])
                .Select(v => NUM(v));
            if (!relatedNumericValues.Any() || relatedNumericValues.All(v => v.GetType() == valueToConvert.GetType()))
                return valueToConvert;

            // DateTime is a special case here - if any of the values (whether the target "o" or any of the related values) is a DateTime
            // then the returned value must be a DateTime. If any of the other values exceed VBScript's expressible date range then an
            // Overflow error occurs. (Most other types can be expanded - if the was an Int16 but there are related values that exceed
            // Int16's range then a different type will be returned).
            //   eg. For i = Date() To 1000000
            //   eg. For i = 1 To Date()
            if ((valueToConvert is DateTime) || relatedNumericValues.Any(v => v is DateTime))
            {
                var relatedNonDateNumericValuesAsDouble = relatedNumericValues
                    .Where(v => !(v is DateTime)) // Ignore Dates since they obviously can't cause a Date overflow!
                    .Select(v => Convert.ToDouble(v));
                if (relatedNonDateNumericValuesAsDouble.Any())
                {
                    if (relatedNonDateNumericValuesAsDouble.Min() < DateToDouble(VBScriptConstants.EarliestPossibleDate))
                        throw new VBScriptOverflowException(relatedNonDateNumericValuesAsDouble.Min());
                    if (relatedNonDateNumericValuesAsDouble.Max() > DateToDouble(VBScriptConstants.LatestPossibleDate))
                        throw new VBScriptOverflowException(relatedNonDateNumericValuesAsDouble.Max());
                }
                return (valueToConvert is DateTime) ? valueToConvert : DoubleToDate(Convert.ToDouble(valueToConvert));
            }

            // If turns out that Decimal is a special case as well.. even when there is a loop constraint that is sufficiently large that
            // it requires a Double and will not fit into the VBScript Currency, the loop variable type remains type Currency and will
            // overflow (note that the relatedNumericValuesAsDouble set defined below will never have to consider dates since if any
            // of the values are dates then we will have exited above, so we don't have to worry about doing a TotalDays calculation
            // rather than calling Convert.ToDouble for everything).
            var relatedNumericValuesAsDouble = relatedNumericValues
                .Select(v => Convert.ToDouble(v));
            if ((valueToConvert is Decimal) || relatedNumericValues.Any(v => v is Decimal))
            {
                if (relatedNumericValues.Any())
                {
                    if (relatedNumericValuesAsDouble.Min() < (double)VBScriptConstants.MinCurrencyValue)
                        throw new VBScriptOverflowException(relatedNumericValuesAsDouble.Min());
                    if (relatedNumericValuesAsDouble.Max() > (double)VBScriptConstants.MaxCurrencyValue)
                        throw new VBScriptOverflowException(relatedNumericValuesAsDouble.Max());
                }
                return (valueToConvert is Decimal) ? valueToConvert : Convert.ToDecimal(valueToConvert);
            }

            // So now we know that the valueToConvert and numericValuesTheTypeMustBeAbleToContain are not all of the same type and that
            // none of them are dates. What we need to do now is take the "biggest" of all of these types and use that.
            return new[] { valueToConvert }.Concat(relatedNumericValues.Select(v => GetAsVBScriptNumber(v)))
                .Select(v =>
                {
                    // Note: Even though double can hold larger numbers then decimal, VBScript prefers decimal - eg. in the loop
                    // "For i = CDbl(0) To CDbl(1) Step CCur(0.2)" the loop variable will be a decimal
                    if (v == null)
                        throw new ArgumentNullException("v");
                    if (v is byte)
                        return Tuple.Create<int, Func<object, object>>(1, value => Convert.ToByte(value));
                    if (v is short)
                        return Tuple.Create<int, Func<object, object>>(2, value => Convert.ToInt16(value));
                    if (v is int)
                        return Tuple.Create<int, Func<object, object>>(3, value => Convert.ToInt32(value));
                    if (v is float)
                        return Tuple.Create<int, Func<object, object>>(4, value => Convert.ToSingle(value));
                    if (v is double)
                        return Tuple.Create<int, Func<object, object>>(5, value => Convert.ToDouble(value));
                    if (v is decimal)
                        return Tuple.Create<int, Func<object, object>>(6, value => Convert.ToDecimal(value));
                    throw new ArgumentException("Unsupported numeric type: " + v.GetType());
                })
                .OrderBy(v => v.Item1)
                .Last()
                .Item2(valueToConvert);
        }

        /// <summary>
        /// This is similar to NullableNUM in that it is used for comparisons involving date literals, where the other side has to be interpreted as
        /// a date but must also support null. It wraps DATE (and so supports all VBScript date parsing methods).
        /// </summary>
        public object NullableDATE(object o)
        {
            o = VAL(o);
            return (o == DBNull.Value) ? DBNull.Value : (object)DATE(o);
        }

        /// <summary>
        /// Reduce a reference down to a date value type, applying VBScript defaults logic and then taking a date representation. Numeric values are
        /// acceptable (taken as the number of days since the zero date, with support for fractional values). String values are acceptable, they will be
        /// parsed using VBScript's rules (and taking into account culture, where applicable). Null is acceptable and will return in the zero date being
        /// returned, DBNull.Value (aka VBScript Null) is not acceptable and will return in an exception being raised.
        /// </summary>
        public DateTime DATE(object o, string optionalExceptionMessageForInvalidContent = null)
        {
            o = VAL(o);
            if (o == null)
                return VBScriptConstants.ZeroDate;
            if (o == DBNull.Value)
                throw new InvalidUseOfNullException(optionalExceptionMessageForInvalidContent);
            if (o is DateTime)
                return (DateTime)o;

            double? numericValue;
            try
            {
                numericValue = Convert.ToDouble(o);
            }
            catch
            {
                numericValue = null;
            }
            try
            {
                if (numericValue != null)
                    return DateParser.Default.Parse(numericValue.Value);
                return DateParser.Default.Parse(o.ToString());
            }
            catch (OverflowException e)
            {
                throw new VBScriptOverflowException(optionalExceptionMessageForInvalidContent, e);
            }
            catch (Exception e)
            {
                throw new TypeMismatchException(optionalExceptionMessageForInvalidContent, e);
            }
        }

        /// <summary>
        /// VBScript only supports a limited set of number types, when NUM is called to determine what type that loop variable should be, it should only be
        /// one of those set. This will ensure that a value is returned as a number that is one of those types (or an exception will be raised, it will
        /// never return null).
        /// </summary>
        private object GetAsVBScriptNumber(object value)
        {
            // Handle the cases of null, DBNull.Value, empty string, booleans, Dates, etc.. (some of these will error, the function will return null if it
            // can't help)
            var specialCaseResult = TryToGetNumberConsideringSpecialCases(value);
            if (specialCaseResult != null)
                return specialCaseResult;

            // Get a value-type reference (we shouldn't get null here, it should be dealt with by TryToGetNumberConsideringSpecialCases - we're going to
            // need to inspect the type of the value here and we can't do that if we get null)
            value = VAL(value);
            if (value == null)
                throw new Exception("Expected TryToGetNumberConsideringSpecialCases to deal with the case of a null value");

            // If the value is a string then that will always become a Double (it doesn't matter if it's "1", which could easily be an "Integer", it will
            // always become a Double)
            if (value is string)
            {
                // Convert.ToDouble seems to match VBScript's string-to-number behaviour pretty well (see the test cases for more details) with one exception;
                // VBScript will tolerate whitespace between a negative sign and the start of the content, so we need to do consider replacements (any "-"
                // followed by whitespace should become just "-")
                return Convert.ToDouble(SpaceFollowingMinusSignRemover.Replace(value.ToString(), "-"));
            }

            // These are the types that map directly onto VBScript types, we can return these unaltered
            if ((value is byte)
            || (value is short)    // aka Int16 aka VBScript "Integer"
            || (value is int)      // aka VBScript "Long"
            || (value is float)    // aka Single aka VBScript "Single"
            || (value is double)   // aka VBScript "Double"
            || (value is decimal)) // aka VBScript "Currency"
                return value;

            // The following are types that we need to manipulate into VBScript-understood types (sbyte, for example, needs bumping into a short / Int16
            // since that is the smallest number that can contain its range of -128 to 127). Note: VBScript has no "long" (Int64) integer type so that goes
            // up to a double. It also has no concept of unsigned types so they need to go into the next bracket.
            if ((value is sbyte) || (value is char))
                return Convert.ToInt16(value);
            else if (value is ushort)
                return Convert.ToInt32(value);
            else if ((value is long) || (value is uint) || (value is ulong))
                return Convert.ToDouble(value);
            else
                throw new ArgumentException("Unsupported type, do not know how to fit it to a VBScript numeric type: " + value.GetType());
        }
        private static Regex SpaceFollowingMinusSignRemover = new Regex(@"-\s+", RegexOptions.Compiled);

        /// <summary>
        /// There are some values which can be treated as numbers (such as Empty, booleans and dates), some of which result in error if this is attempted
        /// (such as blank strings and Null - aka DBNull.Value). This will return a number if a number could be extracted via one of these special cases
        /// and throw an exception if it would not be allowed. If the value is not one of these special cases then null is returned. Note that this will
        /// push the value through the VAL method, so if value is not a value-type reference then it must be reducable to one (otherwise an exception
        /// will be thrown).
        /// </summary>
        private object TryToGetNumberConsideringSpecialCases(object value)
        {
            value = VAL(value);
            if (value == null)
                return (Int16)0; // Return an "Integer" for VBScript Empty
            if (value == DBNull.Value)
                throw new InvalidUseOfNullException();
            if (IsVBScriptNothing(value))
                throw new ObjectVariableNotSetException();
            if ((value as string) == "")
                throw new TypeMismatchException();
            if (value is bool)
                return (bool)value ? (Int16)(-1) : (Int16)0; // Return an "Integer" for True / False
            if (value is DateTime)
                return DateToDouble((DateTime)value);
            return null;
        }

        private double DateToDouble(DateTime value)
        {
            return ((DateTime)value).Subtract(VBScriptConstants.ZeroDate).TotalDays;
        }

        private DateTime DoubleToDate(double value)
        {
            return VBScriptConstants.ZeroDate.AddDays(value);
        }

        /// <summary>
        /// Apply the same logic as STR but allow DBNull.Value (returning it back). This conversion should only used for comparisons with string literals,
        /// where special rules apply (which makes the method slightly less useful than NUM, which is used in comparisons with numeric literals but also
        /// in some other cases, such as FOR loops).
        /// </summary>
        public object NullableSTR(object o)
        {
            o = VAL(o);
            return (o == DBNull.Value) ? DBNull.Value : (object)STR(o);
        }

        /// <summary>
        /// Reduce a reference down to a string value type, applying VBScript defaults logic and then taking a string representation. Null is acceptable
        /// and will return in a blank string being returned, DBNull.Value (aka VBScript Null) is not acceptable and will return in an exception being
        /// raised.
        /// </summary>
        public string STR(object o, string optionalExceptionMessageForInvalidContent = null)
        {
            o = VAL(o, optionalExceptionMessageForInvalidContent);
            if (o == null)
                return "";
            if (o == DBNull.Value)
                throw new InvalidUseOfNullException(optionalExceptionMessageForInvalidContent);
            if (o.GetType().IsArray)
                throw new TypeMismatchException(optionalExceptionMessageForInvalidContent);

            // Dates should be the only data type we have to special-case - booleans, for example, are fine (the casing is consistent between C# and
            // VBScript)
            if (o is DateTime)
                return DateToString((DateTime)o);
            return o.ToString();
        }

        private string DateToString(DateTime value)
        {
            var dateComponent = (value.Date == VBScriptConstants.ZeroDate) ? "" : value.ToShortDateString();
            var timeComponent = ((value.TimeOfDay == TimeSpan.Zero) && (dateComponent != "")) ? "" : value.ToLongTimeString();
            if ((dateComponent != "") && (timeComponent != ""))
                return dateComponent + " " + timeComponent;
            return dateComponent + timeComponent;
        }

        /// <summary>
        /// Layer an enumerable wrapper over a reference, if possible (an exception will be thrown if not)
        /// </summary>
        public IEnumerable ENUMERABLE(object o)
        {
            if (o == null)
                throw new ArgumentNullException("o");

            // VBScript will only consider object references to be enumerable (unlike C#, which will consider a string to be an enumerable set
            // characters, for example)
            if (IsVBScriptValueType(o))
                throw new ArgumentException("Object not a collection");

            // Try casting to IEnumerable first - it's the easiest approach and will work with (many) managed references and some COM object
            var enumerable = o as IEnumerable;
            if (enumerable != null)
                return enumerable;

            // Failing that, attempt access through IDispatch - try calling the method with DispId -4 (if there is one) and casting the return
            // value to IEnumVariant and then wrapping up into a managed IEnumerable
            if (IDispatchAccess.ImplementsIDispatch(o))
            {
                object enumerator;
                try
                {
                    enumerator = IDispatchAccess.Invoke<object>(o, IDispatchAccess.InvokeFlags.DISPATCH_METHOD, -4);
                }
                catch (MissingMemberException)
                {
                    throw new ArgumentException("IDispatch reference does not have a method with DispId -4");
                }
                var enumeratorAsEnumVariant = enumerator as IEnumVariant;
                if (enumeratorAsEnumVariant == null)
                    throw new ArgumentException("IDispatch reference has a DispId -4 return value that does not implement IEnumVariant");
                return new ManagedEnumeratorWrapper(new IDispatchEnumeratorWrapper(enumeratorAsEnumVariant));
            }

            // Give up and throw the VBScript error message
            throw new ArgumentException("Object not a collection");
        }

        /// <summary>
        /// Reduce a reference down to a boolean, throwing an exception if this is not possible. This will apply the same logic as VAL but then
        /// require a numeric value or null, otherwise an exception will be raised. Zero and null equate to false, non-zero numbers to true.
        /// </summary>
        public bool IF(object o)
        {
            if (o == null)
                return false;

            // The BOOL function will not accept VBScript Null since it can't strictly be translated into a boolean, in the context of an IF
            // condition it can be taken to mean false, though
            var value = VAL(o);
            if (value == DBNull.Value)
                return false;
            return BOOL(value);
        }

        /// <summary>
        /// This is used to wrap arguments such that those that must be passed ByVal can have changes reflected after a method call completes
        /// </summary>
        public IBuildCallArgumentProviders ARGS
        {
            get { return new DefaultCallArgumentProvider(this); }
        }

        /// <summary>
        /// This requires a target with optional member accessors and arguments - eg. "Test" is a target only, "a.Test" has target "a" with one
        /// named member "Test", "a.Test(0)" has target "a", named member "Test" and a single argument "0". The expression "a(Test(0))" would
        /// require nested CALL executions, one with target "Test" and a single argument "0" and a second with target "a" and a single
        /// argument which was the result of the first call.
        /// </summary>
        public object CALL(object target, IEnumerable<string> members, IProvideCallArguments argumentProvider)
        {
            if (members == null)
                throw new ArgumentNullException("members");
            if (argumentProvider == null)
                throw new ArgumentNullException("argumentProvider");

            // Note: The arguments are evaluated here before the CALL is attempted - so if the target or members are invalid then this is only
            // determined AFTER processing the arguments. This is correct behaviour (consistent with VBScript).
            var arguments = argumentProvider.GetInitialValues().ToArray();
            try
            {
                return CALL(target, members, arguments, argumentProvider.UseBracketsWhereZeroArguments);
            }
            finally
            {
                // Even if an exception if thrown somewhere in the target call, any ByRef argument values that were changed must be persisted
                for (var index = 0; index < arguments.Length; index++)
                    argumentProvider.OverwriteValueIfByRef(index, arguments[index]);
            }
        }

        /// <summary>
        /// Note: The arguments array elements may be mutated if the call target has "ref" method arguments.
        /// </summary>
        private object CALL(object target, IEnumerable<string> members, object[] arguments, bool useBracketsWhereZeroArguments)
        {
            if (members == null)
                throw new ArgumentNullException("members");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            var memberAccessorsArray = members.ToArray();
            if (memberAccessorsArray.Any(m => string.IsNullOrWhiteSpace(m)))
                throw new ArgumentException("Null/blank value in members set");

            // If the target is a value type then neither member accessors nor arguments are valid (eg. "#1 1#()" or "Null.ToString()"), so check
            // for those immediately. If we don't check then then might actually be allowed since C# treats everything as an object and everything
            // has a ToString method. Note that some value types are compile errors (eg. "123.ToString()") and so those should have been caught by
            // the translation process. VBScript considers arrays to be value types (ie. when returning an array from a function, you don't have
            // to call SET), so we need to exclude that case since clearly arguments CAN be used with arrays.
            var isArray = (target != null) && target.GetType().IsArray;
            if (!isArray && IsVBScriptValueType(target))
            {
                string targetDescription;
                if (target == null)
                    targetDescription = "[undefined]"; // This is what VBScript shows for "Empty.ToString()"
                else if (target == DBNull.Value)
                    targetDescription = "Null";
                else
                    targetDescription = STR(target);
                if (memberAccessorsArray.Any())
                    throw new ObjectRequiredException("'" + targetDescription + "'");
                if (arguments.Any() || useBracketsWhereZeroArguments)
                    throw new TypeMismatchException("'" + targetDescription + "'");
            }

            // Deal with special case of a delegate first (as of May 2015, there should't be any way for one of these to sneak in here, but if
            // GetRef gets implemented in the future then that may change). It's not valid for there to be any member accessors, the delegate
            // must either be a direct reference, meaning its being passed around as a function pointer or sorts, or it must be an execution
            // of the delegate (indicated by the presence of arguments or by there being brackets following the reference, even when there
            // are zero arguments).
            var delegateTarget = target as Delegate;
            if (delegateTarget != null)
            {
                if (memberAccessorsArray.Any())
                    throw new ArgumentException("May not specify any member accessors when target is a delegate");
                if (arguments.Any() || useBracketsWhereZeroArguments)
                    return delegateTarget.DynamicInvoke(arguments);
                return delegateTarget;
            }

            // If there are no member accessor, no arguments and no zero-argument brackets then just return the object directly (if the caller
            // wants this to be a value type then they'll have to use VAL to try to apply that requirement). This is the only case in which
            // it's acceptable for target to be null.
            // - Frankly, I'm not sure why target can EVER be null, even if there ARE no member accessor or aguments... unfortunately, there
            //   are currently no tests illustrating why at this time and the code here explicitly allows for a null target (in this case),
            //   so I'm going to leave this be for now (May 2015).
            var noMemberAccessorsOrArguments = !memberAccessorsArray.Any() && !arguments.Any();
            if (noMemberAccessorsOrArguments && !useBracketsWhereZeroArguments)
                return target;
            else if (noMemberAccessorsOrArguments && useBracketsWhereZeroArguments && (target != null))
            {
                // If "a" is an array then "a()" will always throw a "Subscript out of range" exception
                if (isArray)
                    throw new SubscriptOutOfRangeException();
                throw new TypeMismatchException(); // It's not an array and it's not a function, must be a type mismatch
            }
            else if (target == null)
                throw new TypeMismatchException("The target reference may only be null if there are no member accessors or arguments specified (and no brackets in the zero-argument case)");

            // If there are no member accessors but there are arguments then
            // 1. Attempt array access (this is the only time that array access is acceptable (o.Names(0) is not allowed by VBScript if the
            //    "Names" property is an array, it will only work if Names if a Function or indexed Property)
            // 2. If target implements IDispatch then try accessing the DispId zero method or property, passing the arguments
            // 3. If it's not an IDispatch reference, then the arguments will be passed to a method or indexed property that will accept them,
            //    taking into account the IsDefault attribute
            if (!memberAccessorsArray.Any() && arguments.Any())
                return InvokeGetter(target, null, arguments, onlyConsiderMethods: false);

            // If there are member accessors but no arguments then we walk down each member accessor, no defaults are considered
            // - If useBracketsWhereZeroArguments is true then only consider methods in this lookup (eg. "a.Name()" in VBScript requires that "Name"
            //   be a method.. when calling external code, such as COM components, at least; internally, VBScript considers property getters to be
            //   functions so "a.Name" and "a.Name()" will both retrieve the value of a public get-able VBScript class property, but if "a" is an
            //   IDispatch reference, then only IDispatch methods will be acceptable for use). This is the only time where we need to enforce the
            //   "onlyConsiderMethods" option since it is only when there are zero arguments and the absence or presence of brackets that different
            //   logic is required (so other calls to WalkMemberAccessors / InvokeGetter leave that option as false).
            if (!arguments.Any())
                return WalkMemberAccessors(target, memberAccessorsArray, onlyConsiderMethods: useBracketsWhereZeroArguments);

            // If there member accessors AND arguments then all-but-the-last member accessors should be walked through as argument-less lookups
            // and the final member accessor should be a method or property call whose name matches the member accessor. Note that the arguments
            // can never be for an array look up at this point because there are member accessors, therefor it must be a function or property.
            target = WalkMemberAccessors(target, memberAccessorsArray.Take(memberAccessorsArray.Length - 1), onlyConsiderMethods: false);
            var finalMemberAccessor = memberAccessorsArray[memberAccessorsArray.Length - 1];
            if (target == null)
                throw new ArgumentException("Unable to access member \"" + finalMemberAccessor + "\" on null reference");
            return InvokeGetter(target, finalMemberAccessor, arguments, onlyConsiderMethods: false);
        }

        /// <summary>
        /// This will throw an exception for null target or arguments references or if the setting fails (eg. invalid number of arguments,
        /// invalid member accessor - if specified - argument thrown by the target setter). This must not be called with a target reference
        /// only (null optionalMemberAccessor and zero arguments) as it would need to change the caller's reference to target, which is not
        /// possible (in that case, a straight assignment should be generated - no call to SET required). Note that the valueToSetTo argument
        /// comes before any others since VBScript will evaulate the right-hand side of the assignment before the left, which may be important
        /// if an error is raised at some point in the operation.
        /// </summary>
        public void SET(object valueToSetTo, object target, string optionalMemberAccessor, IProvideCallArguments argumentProvider)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (argumentProvider == null)
                throw new ArgumentNullException("argumentProvider");

            var arguments = argumentProvider.GetInitialValues().ToArray();
            if ((optionalMemberAccessor == null) && !arguments.Any())
                throw new ArgumentException("This must be called with a non-null optionalMemberAccessor and/or one or more arguments, null optionalMemberAccessor and zero arguments is not supported");

            var cacheKey = new InvokerCacheKey(target.GetType(), optionalMemberAccessor, arguments.Length, onlyConsiderMethods: false);
            SetInvoker invoker;
            if (!_setInvokerCache.TryGetValue(cacheKey, out invoker))
            {
                invoker = GenerateSetInvoker(target, optionalMemberAccessor, arguments);
                _setInvokerCache.TryAdd(cacheKey, invoker);
            }
            invoker(target, arguments, valueToSetTo);
            for (var index = 0; index < arguments.Length; index++)
                argumentProvider.OverwriteValueIfByRef(index, arguments[index]);
        }

        /// <summary>
        /// The arguments set must be an array since its contents may be mutated if the call target has "ref" parameters
        /// </summary>
        private object InvokeGetter(object target, string optionalName, object[] arguments, bool onlyConsiderMethods)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            var cacheKey = new InvokerCacheKey(target.GetType(), optionalName, arguments.Length, onlyConsiderMethods);
            GetInvoker invoker;
            if (!_getInvokerCache.TryGetValue(cacheKey, out invoker))
            {
                invoker = GenerateGetInvoker(target, optionalName, arguments, onlyConsiderMethods);
                _getInvokerCache.TryAdd(cacheKey, invoker);
            }
            return invoker(target, arguments);
        }

        private delegate object GetInvoker(object target, object[] arguments);
        private GetInvoker GenerateGetInvoker(object target, string optionalName, IEnumerable<object> arguments, bool onlyConsiderMethods)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            var argumentsArray = arguments.ToArray();
            var targetType = target.GetType();
            if (targetType.IsArray)
            {
                if (targetType.GetArrayRank() != argumentsArray.Length)
                    throw new SubscriptOutOfRangeException("Argument count (" + argumentsArray.Length + ") does not match array rank (" + targetType.GetArrayRank() + ")");

                var arrayTargetParameter = Expression.Parameter(typeof(object), "target");
                var indexesParameter = Expression.Parameter(typeof(object[]), "arguments");
                var arrayAccessExceptionParameter = Expression.Parameter(typeof(Exception), "e");
                return Expression.Lambda<GetInvoker>(
                    Expression.TryCatch(
                        Expression.Convert(
                            Expression.ArrayAccess(
                                Expression.Convert(arrayTargetParameter, targetType),
                                Enumerable.Range(0, argumentsArray.Length).Select(index =>
                                    GetVBScriptStyleArrayIndexParsingExpression(
								        Expression.ArrayAccess(indexesParameter, Expression.Constant(index))
                                    )
                                )
                            ),
                            typeof(object) // Without this we may get an "Expression of type 'System.Int32' cannot be used for return type 'System.Object'" or similar
                        ),
                        Expression.Catch(
                            arrayAccessExceptionParameter,
                            Expression.Throw(
                                GetNewArgumentException(
                                    "Error accessing array with specified indexes (likely a non-numeric array index/argument or index-out-of-bounds)",
                                    arrayAccessExceptionParameter
                                ),
                                typeof(object)
                            )
                        )
                    ),
                    new[]
			        {
				        arrayTargetParameter,
				        indexesParameter
			        }
                ).Compile();
            }

            var errorMessageMemberDescription = (optionalName == null) ? "default member" : ("member \"" + optionalName + "\"");
            if (onlyConsiderMethods)
                errorMessageMemberDescription += " (allowing methods only)";
            if (argumentsArray.Length == 0)
                errorMessageMemberDescription = "parameter-less " + errorMessageMemberDescription;
            else
                errorMessageMemberDescription += " that will accept " + argumentsArray.Length + " argument(s)";

            if (IDispatchAccess.ImplementsIDispatch(target))
            {
                int dispId;
                if (optionalName == null)
                    dispId = 0;
                else
                {
                    try
                    {
                        // We don't use the nameRewriter here since we won't have rewritten the COM component, it's the C# generated from the
                        // VBScript source that we may have rewritten
                        dispId = IDispatchAccess.GetDispId(target, optionalName);
                    }
                    catch (Exception e)
                    {
                        throw new ObjectVariableNotSetException("Unable to identify " + errorMessageMemberDescription + " (target implements IDispatch)", e);
                    }
                }
                return (invokeTarget, invokeArguments) =>
                {
                    try
                    {
                        return IDispatchAccess.Invoke<object>(
                            invokeTarget,
                            onlyConsiderMethods
                                ? IDispatchAccess.InvokeFlags.DISPATCH_METHOD 
                                : IDispatchAccess.InvokeFlags.DISPATCH_METHOD | IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYGET,
                            dispId,
                            invokeArguments.ToArray()
                        );
                    }
                    catch (Exception e)
                    {
                        throw new ArgumentException("Error executing " + errorMessageMemberDescription + " (target implements IDispatch): " + e.GetBaseException(), e);
                    }
                };
            }

            var possibleMethods = (optionalName == null)
                ? GetDefaultGetMethods(targetType, argumentsArray.Length)
                : GetNamedGetMethods(targetType, optionalName, argumentsArray.Length);
            if (!possibleMethods.Any())
                throw new ObjectVariableNotSetException("Unable to identify " + errorMessageMemberDescription);

            var method = possibleMethods.First();
            var targetParameter = Expression.Parameter(typeof(object), "target");
            var argumentsParameter = Expression.Parameter(typeof(object[]), "arguments");
            var exceptionParameter = Expression.Parameter(typeof(Exception), "e");

            // If there are any ByRef arguments then we need local variables that are of ByRef types. These will be populated with values from the
            // arguments array before the method call, then used AS arguments in the method call, then their values copied back into the slots in
            // the arguments array that they came from.
            var methodParameters = method.GetParameters();
            Expression[] argExpressions;
            IEnumerable<Expression> byRefArgAssignmentsForMethodCall, byRefArgAssignmentsForReturn;
            IEnumerable<ParameterExpression> variablesToDeclareForHandlingOfArguments;
            if ((methodParameters.Length == 1) && ParameterIsObjectParamsArray(methodParameters[0]))
            {
                // Add support for the rudimentary params support considered by the "GetGetMethods" function - if the target function has a single
                // argument of type params object[] then pass the arguments array straight in for that argument (no need to try to break it down
                // into individual entries for multiple arguments, since there is only a single argument!). C# does not support "ref params" and
                // so we won't worry about passing "by ref" here.
                argExpressions = new[] { argumentsParameter };
                byRefArgAssignmentsForMethodCall = new Expression[0];
                byRefArgAssignmentsForReturn = new Expression[0];
                variablesToDeclareForHandlingOfArguments = new ParameterExpression[0];
            }
            else
            {
                var argParameters = methodParameters.Select((p, index) => new { Parameter = p, Index = index });
                argExpressions = argParameters.Select(a =>
                    a.Parameter.ParameterType.IsByRef
                        ? (Expression)Expression.Variable(
                            GetNonByRefType(a.Parameter.ParameterType),
                            a.Parameter.Name
                        )
                        : Expression.Convert(
                            Expression.ArrayAccess(argumentsParameter, Expression.Constant(a.Index)),
                            a.Parameter.ParameterType
                        )
                    )
                    .ToArray();
                byRefArgAssignmentsForMethodCall = argExpressions
                    .Select((a, index) => new { Parameter = a, Index = index })
                    .Where(a => a.Parameter is ParameterExpression)
                    .Select(a =>
                        Expression.Assign(
                            a.Parameter,
                            Expression.Convert(
                                Expression.ArrayAccess(argumentsParameter, Expression.Constant(a.Index)),
                                a.Parameter.Type
                            )
                        )
                    );
                byRefArgAssignmentsForReturn = argExpressions
                    .Select((a, index) => new { Parameter = a, Index = index })
                    .Where(a => a.Parameter is ParameterExpression)
                    .Select(a =>
                        Expression.Assign(
                            Expression.ArrayAccess(argumentsParameter, Expression.Constant(a.Index)),
                            a.Parameter
                        )
                    );
                variablesToDeclareForHandlingOfArguments = argExpressions.OfType<ParameterExpression>();
            }

            var methodCall = Expression.Call(
                Expression.Convert(targetParameter, targetType),
                method,
                argExpressions
            );

            Expression[] methodCallAndAndResultAssignments;
            var resultVariable = Expression.Variable(typeof(object));
            if (method.ReturnType == typeof(void))
            {
                // If the method has no return type then we'll need to just return null since the GetInvoker delegate that
                // we're trying to compile to has a return type of "object" (so execute methodCall and then set the result
                // variable to null)
                methodCallAndAndResultAssignments = new Expression[]
                {
                    methodCall,
                    Expression.Assign(
                        resultVariable,
                        Expression.Constant(null, typeof(object))
                    )
                };
            }
            else
            {
                // If there IS a a return type then assign the method call's return value to the result variable
                methodCallAndAndResultAssignments = new Expression[]
                {
                    Expression.Assign(
                        resultVariable,
                        methodCall.Type.IsValueType ? (Expression)Expression.Convert(methodCall, typeof(object)) : methodCall
                    )
                };
            }

            // Note: The Throw expression will require a return type to be specified since the Try block has a return type
            // - without this a runtime "Body of catch must have the same type as body of try" exception will be raised
            Expression executionException = Expression.Throw(
                GetNewArgumentException("Error executing " + errorMessageMemberDescription, exceptionParameter),
                typeof(object)
            );
            Expression exceptionHandlingMethodCallExpression;
            if (byRefArgAssignmentsForReturn.Any())
            {
                // If there are ByRef arguments that need setting then this must be done in a finally block, since they must
                // be set even if the target method throws an exception at some point. If there are no ByRef arguments then
                // don't generate a try..catch..finall block at all since there will be an exception thrown about a block
                // for the finally that is an empty expression collection.
                exceptionHandlingMethodCallExpression = Expression.TryCatchFinally(
                    Expression.Block(
                        methodCallAndAndResultAssignments.Concat(new[] { resultVariable })
                    ),
                    Expression.Block(byRefArgAssignmentsForReturn), // Finally block (always set any ByRef arguments, even if an exception is thrown)
                    Expression.Catch(exceptionParameter, executionException)
                );
            }
            else
            {
                exceptionHandlingMethodCallExpression = Expression.TryCatch(
                    Expression.Block(
                        methodCallAndAndResultAssignments.Concat(new[] { resultVariable })
                    ),
                    Expression.Catch(exceptionParameter, executionException)
                );
            }
            var variablesToDeclare = variablesToDeclareForHandlingOfArguments.Concat(new[] { resultVariable });
            return Expression.Lambda<GetInvoker>(
                Expression.Block(
                    variablesToDeclare,
                    Expression.Block(
                        byRefArgAssignmentsForMethodCall.Cast<Expression>().Concat(new[] { exceptionHandlingMethodCallExpression })
                    )
                ),
                targetParameter,
                argumentsParameter
            ).Compile();
        }

        private delegate void SetInvoker(object target, object[] arguments, object value);
        private SetInvoker GenerateSetInvoker(object target, string optionalMemberAccessor, IEnumerable<object> arguments)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            // If there are no member accessors but there are arguments then firstly attempt array access, this is the only time that array access is acceptable
            // (o.Names(0) is not allowed by VBScript if the "Names" property is an array, it will only work if Names is a Function or indexed Property and since
            // we're attempting a set here, it will only work if it's an indexed Property)
            var argumentsArray = arguments.ToArray();
            var targetType = target.GetType();
            if ((optionalMemberAccessor == null) && targetType.IsArray)
            {
                if (targetType.GetArrayRank() != argumentsArray.Length)
                    throw new ArgumentException("Argument count (" + argumentsArray.Length + ") does not match array rank (" + targetType.GetArrayRank() + ")");

                // Without the targetType.GetElementType() specified for the Expression.Catch, a "Body of catch must have the same type as body of try" exception
                // will be raised. Since I would have expected the try block to return nothing (void) I'm not quite sure why this is.
                var arrayTargetParameter = Expression.Parameter(typeof(object), "target");
                var indexesParameter = Expression.Parameter(typeof(object[]), "arguments");
                var arrayAccessExceptionParameter = Expression.Parameter(typeof(Exception), "e");
                var arrayValueParameter = Expression.Parameter(typeof(object), "value");
                return Expression.Lambda<SetInvoker>(
                    Expression.TryCatch(
                        Expression.Assign(
                            Expression.ArrayAccess(
                                Expression.Convert(arrayTargetParameter, targetType),
                                Enumerable.Range(0, argumentsArray.Length).Select(index =>
                                    GetVBScriptStyleArrayIndexParsingExpression(
                                        Expression.ArrayAccess(indexesParameter, Expression.Constant(index))
                                    )
                                )
                            ),
                            Expression.Convert(
                                arrayValueParameter,
                                targetType.GetElementType()
                            )
                        ),
                        Expression.Catch(
                            arrayAccessExceptionParameter,
                            Expression.Throw(
                                GetNewArgumentException(
                                    "Error accessing array with specified indexes (likely a non-numeric array index/argument or index-out-of-bounds)",
                                    arrayAccessExceptionParameter
                                ),
                                targetType.GetElementType()
                            )
                        )
                    ),
                    new[]
			        {
				        arrayTargetParameter,
				        indexesParameter,
                        arrayValueParameter
			        }
                ).Compile();
            }

            var errorMessageMemberDescription = (optionalMemberAccessor == null) ? "default member" : ("member \"" + optionalMemberAccessor + "\"");
            if (argumentsArray.Length == 0)
                errorMessageMemberDescription = "parameter-less " + errorMessageMemberDescription;
            else
                errorMessageMemberDescription += " that will accept " + argumentsArray.Length + " argument(s)";

            if (IDispatchAccess.ImplementsIDispatch(target))
            {
                int dispId;
                if (optionalMemberAccessor == null)
                    dispId = 0;
                else
                {
                    try
                    {
                        // We don't use the nameRewriter here since we won't have rewritten the COM component, it's the C# generated from the
                        // VBScript source that we may have rewritten
                        dispId = IDispatchAccess.GetDispId(target, optionalMemberAccessor);
                    }
                    catch (Exception e)
                    {
                        throw new ArgumentException("Unable to identify " + errorMessageMemberDescription + " (target implements IDispatch)", e);
                    }
                }
                return (invokeTarget, invokeArguments, value) =>
                {
                    try
                    {
                        IDispatchAccess.Invoke<object>(
                            target,
                            IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYPUT,
                            dispId,
                            argumentsArray.Concat(new[] { value }).ToArray()
                        );
                    }
                    catch (Exception e)
                    {
                        throw new ArgumentException("Error executing " + errorMessageMemberDescription + " (target implements IDispatch): " + e.GetBaseException(), e);
                    }
                };
            }

            MethodInfo method;
            if (optionalMemberAccessor != null)
            {
                // If there is a non-null optionalMemberAccessor but no arguments then try setting the non-indexed property, no defaults considered.
                // If there is a non-null optionalMemberAccessor and there are arguments then no defaults are considered and the member accessor
                // must be an indexed property, it can not be an array (see note above about the only place that array access is permitted).
                method = GetNamedSetMethods(targetType, optionalMemberAccessor, argumentsArray.Length).FirstOrDefault();
            }
            else
            {
                // Try accessing a default indexed property (either a a native C# property or a method with IsDefault and TranslatedProperty attributes)
                method = GetDefaultSetMethods(targetType, argumentsArray.Length).FirstOrDefault();
            }
            if (method == null)
                throw new ArgumentException("Unable to identify " + errorMessageMemberDescription);

            var targetParameter = Expression.Parameter(typeof(object), "target");
            var argumentsParameter = Expression.Parameter(typeof(object[]), "arguments");
            var exceptionParameter = Expression.Parameter(typeof(Exception), "e");
            var valueParameter = Expression.Parameter(typeof(object), "value");
            return Expression.Lambda<SetInvoker>(

                Expression.TryCatch(

                    Expression.Call(
                        Expression.Convert(targetParameter, targetType),
                        method,
                        method.GetParameters().Select((arg, index) =>
                        {
                            // The method will have one more parameter than there are arguments since the last parameter is the value to set to
                            return Expression.Convert(
                                (index == argumentsArray.Length)
                                    ? (Expression)valueParameter
                                    : Expression.ArrayAccess(argumentsParameter, Expression.Constant(index)),
                                arg.ParameterType
                            );
                        })
                    ),

                    Expression.Catch(
                        exceptionParameter,
                        Expression.Throw(
                            GetNewArgumentException("Error executing " + errorMessageMemberDescription, exceptionParameter)
                        )
                    )

                ),
                new[]
				{
					targetParameter,
					argumentsParameter,
                    valueParameter
				}
            ).Compile();
        }

        /// <summary>
        /// VBScript expects integer array index values but will accept fractional values or even strings representing integers or fractional values.
        /// Any fractional values are rounded to the closest even number (so 1.5 rounds to 1, 2.5 also rounds to 2, 3.5 rounds to 4, etc..)
        /// </summary>
        private Expression GetVBScriptStyleArrayIndexParsingExpression(Expression index)
        {
            if (index == null)
                throw new ArgumentNullException("index");

            var returnValueParameter = Expression.Parameter(typeof(int), "retVal");
            return Expression.Block(
                typeof(int),
                new[] { returnValueParameter },
                Expression.IfThenElse(
                    Expression.TypeIs(index, typeof(int)),
                    Expression.Assign(
                        returnValueParameter,
                        Expression.Convert(
                            index,
                            typeof(int)
                        )
                    ),
                    Expression.IfThenElse(
                        Expression.TypeIs(index, typeof(double)),
                        Expression.Assign(
                            returnValueParameter,
                            Expression.Convert(
                                Expression.Call(
                                    null,
                                    typeof(Math).GetMethod("Round", new[] { typeof(double), typeof(MidpointRounding) }),
                                    Expression.Convert(
                                        index,
                                        typeof(double)
                                    ),
                                    Expression.Constant(MidpointRounding.ToEven)
                                ),
                                typeof(int)
                            )
                        ),
                        Expression.Assign(
                            returnValueParameter,
                            Expression.Convert(
                                Expression.Call(
                                    null,
                                    typeof(Math).GetMethod("Round", new[] { typeof(double), typeof(MidpointRounding) }),
                                    Expression.Call(
                                        null,
                                        typeof(double).GetMethod("Parse", new[] { typeof(string) }),
                                        Expression.Call(
                                            index,
                                            typeof(object).GetMethod("ToString")
                                        )
                                    ),
                                    Expression.Constant(MidpointRounding.ToEven)
                                ),
                                typeof(int)
                            )
                        )
                    )
                ),
                returnValueParameter
            );
        }

        private Expression GetNewArgumentException(string message, ParameterExpression exceptionParameter)
        {
            if (string.IsNullOrWhiteSpace(message))
                throw new ArgumentException("Null/blank message specified");
            if (exceptionParameter == null)
                throw new ArgumentNullException("exceptionParameter");

            return Expression.New(
                typeof(ArgumentException).GetConstructor(new[] { typeof(string), typeof(Exception) }),
                Expression.Call(
                    typeof(string).GetMethod("Concat", new[] { typeof(object), typeof(object) }),
                    Expression.Constant(message.Trim() + ": "),
                    Expression.Call(
                        exceptionParameter,
                        typeof(Exception).GetMethod("GetBaseException")
                    )
                ),
                exceptionParameter
            );
        }

        private Type GetNonByRefType(Type byRefType)
        {
            if (byRefType == null)
                throw new ArgumentNullException("byRefType");
            if (!byRefType.IsByRef)
                throw new ArgumentException("Must be a type which reports IsByRef true");

            var byRefFullName = byRefType.FullName;
            var nonByRefFullName = byRefFullName.Substring(0, byRefFullName.Length - 1);
            return Type.GetType(
                byRefType.AssemblyQualifiedName.Replace(byRefFullName, nonByRefFullName),
                true
            );
        }

        private int ApplyVBScriptIndexLogicToNonIntegerValue(double value)
        {
            return (int)Math.Round(value, MidpointRounding.ToEven); // This is what effectively what VBScript does
        }

        private object WalkMemberAccessors(object target, IEnumerable<string> memberAccessors, bool onlyConsiderMethods)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (memberAccessors == null)
                throw new ArgumentNullException("memberAccessors");

            foreach (var memberAccessor in memberAccessors)
            {
                if (string.IsNullOrWhiteSpace(memberAccessor))
                    throw new ArgumentException("Encountered null/blank in memberAccessors set");

                // This should only be the case after we've gone round the loop at least once
                if (target == null)
                    throw new ArgumentException("Unable to access member \"" + memberAccessor + "\" on null reference");

                target = InvokeGetter(target, memberAccessor, new object[0], onlyConsiderMethods);
            }
            return target;
        }

        private IEnumerable<MethodInfo> GetDefaultGetMethods(Type type, int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            return GetGetMethods(type, null, DefaultMemberBehaviourOptions.MustBeDefault, MemberNameMatchBehaviourOptions.Precise, numberOfArguments);
        }

        private IEnumerable<MethodInfo> GetDefaultSetMethods(Type type, int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            return GetSetMethods(type, null, DefaultMemberBehaviourOptions.MustBeDefault, MemberNameMatchBehaviourOptions.Precise, numberOfArguments);
        }

        private IEnumerable<MethodInfo> GetNamedGetMethods(Type type, string name, int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Null/blank name specified");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            // There the nameRewriter WILL be considered in case it's trying to access classes we've translated. However, there's also a chance that
            // we could be accessing a non-IDispatch CLR type from somewhere, so GetNamedGetMethods will try to match using the nameRewriter first and
            // then fallback to a perfect match non-rewritten name and finally to a case-insensitive match to a non-rewritten name.
            return GetGetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.UseNameRewriter, numberOfArguments)
                .Concat(
                    GetGetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.Precise, numberOfArguments)
                )
                .Concat(
                    GetGetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.CaseInsensitive, numberOfArguments)
                );
        }

        private IEnumerable<MethodInfo> GetNamedSetMethods(Type type, string name, int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Null/blank name specified");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            // There the nameRewriter WILL be considered in case it's trying to access classes we've translated. However, there's also a chance that
            // we could be accessing a non-IDispatch CLR type from somewhere, so GetNamedGetMethods will try to match using the nameRewriter first and
            // then fallback to a perfect match non-rewritten name and finally to a case-insensitive match to a non-rewritten name.
            return GetSetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.UseNameRewriter, numberOfArguments)
                .Concat(
                    GetSetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.Precise, numberOfArguments)
                )
                .Concat(
                    GetSetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.CaseInsensitive, numberOfArguments)
                );
        }

        private IEnumerable<MethodInfo> GetGetMethods(
            Type type,
            string optionalName,
            DefaultMemberBehaviourOptions defaultMemberBehaviour,
            MemberNameMatchBehaviourOptions memberNameMatchBehaviour,
            int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (!Enum.IsDefined(typeof(DefaultMemberBehaviourOptions), defaultMemberBehaviour))
                throw new ArgumentOutOfRangeException("defaultMemberBehaviour");
            if (!Enum.IsDefined(typeof(MemberNameMatchBehaviourOptions), memberNameMatchBehaviour))
                throw new ArgumentOutOfRangeException("memberNameMatchBehaviour");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            // There is some crazy behaviour around default member accesses when VBScript and COM are combined, which need to be reflected here. For
            // example, if C# class is wrapped up as a COM object and consumed by VBScript and that class is decorated with ComVisible(true) and an
            // unnamed parameter-less member access is made against that object - eg.
            //
            //   Set o = CreateObject("My.ComComponent")
            //   If (o = "Test") Then
            //    ' Something
            //   End If
            // 
            // then various options are considered. If there is a method or property with DispId(0) then that is an obvious place to start, but an
            // error will be raised if has arguments, since zero arguments are being passed in. Also, if multiple members are marked with DispId(0)
            // then it doesn't know which to choose and they are all ignored. In this case, or the case of no DispId(0) members, there are further
            // alternatives considered: If there is a method or property called "Value" with zero arguments then this will queried. If this is not
            // available then the parameter-less "ToString" method will be called - but only if the type or one of its explicit base types declares
            // ComVisible(true). All classes implicitly inherit the base type Object, which declares ComVisible(true), but if that is the only type
            // in the inheritance tree that does then the "ToString" fallback is not considered. The "Value" property / method fallback is acceptable
            // if it is on a type with ComVisible(false) so long as elsewhere in the inheritance tree there is a type other than Object which specifies
            // ComVisible(true). This magic is actually something that happens when the class is exposed as COM and introduced by the C# compiler, it
            // will set the ToString method to have DispId zero if nothing else does and if DispId(0) attributes are found for multiple methods (any-
            // where in a type's inheritance tree) then it will also force the ToString method to have DispId zero (this is talked about in the book
            // ".NET and COM: The Complete Interoperability Guide", though I'm yet to find a definitive reference for the default "Value" member
            // assignment). Depending upon how the component is exposed when consumed through .net, none of this may be required - if we get true
            // back when we call IDispatchAccess.ImplementsIDispatch(target) then IDispatch will be used to query the object and this code path
            // will not be visited. This reflection-based approach is much more complex, but it allows us to cache a lambda for this particular
            // member access on this target type, which means that we'll get some performance benefits if it's called many times. I fully suspect
            // that, at this time, there are various scenarios which are not being handled correctly - I'm trying to cover the obvious cases that
            // VBScript code would have to deal with and need to read ".NET and COM: The Complete Interoperability Guide" more thoroughly and try
            // to apply all of that knowledge when I absorb it!
            //
            // One final thing to consider: these special cases should not apply to classes that were translated from VBScript source since these edge
            // cases are something applied by the C# compiler and so would not have been applied to the original VBScript classes (if a VBScript class
            // has no default method or property and is accessed in a manner as illustrated above, where the object reference needs to be processed as
            // a value reference, then a Type Mismatch error would be raised - a parameter-less "Value" property would be no help if it wasn't explicitly
            // declared to be default). Translated classes will have the "SourceClassName" attribute specifiede on them, which is what the following code
            // uses to differentiate between translated classes and everything else.

            // Try to locate methods and/or properties that meet the criteria; have the correct number of arguments and match the name or are default, if
            // default member access is desired (dealing with the differing logic for translated-from-VBScript types and non-translated-from-VBScript).
            var nameMatcher = (optionalName != null) ? GetNameMatcher(optionalName, memberNameMatchBehaviour) : (name => true);
            var typeWasTranslatedFromVBScript = TypeIsTranslatedFromVBScript(type);
            var typeIsComVisible = TypeIsComVisible(type);
            var typeHasAnyDispIdZeroMember = AnyDispIdZeroMemberExists(type);
            var typeHasAmbiguousDispIdZeroMember = typeHasAnyDispIdZeroMember && DispIdZeroIsAmbiguous(type);
            var allMethods = GetMethodsThatAreNotRelatedToProperties(type);
            Predicate<MethodInfo> matchesArgumentCount = m =>
            {
                // Add some crude support for params array arguments - only dealing with the case where the target method has a single argument of type
                // params object[] (there is corresponding code the in the GenerateGetInvoker for dealing with methods of this form)
                var args = m.GetParameters();
                if (args.Length == numberOfArguments)
                    return true;
                return (args.Length == 1) && ParameterIsObjectParamsArray(args[0]);
            };
            var applicableMethods = allMethods
                    .Where(m => nameMatcher(m.Name))
                    .Where(m =>
                        (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                        (m.GetCustomAttributes(typeof(IsDefault), true).Any()) ||
                        (!typeWasTranslatedFromVBScript && typeIsComVisible && !typeHasAmbiguousDispIdZeroMember && MemberHasDispIdZero(m))
                    )
                    .Where(m => matchesArgumentCount(m));
            var readableProperties = type.GetProperties().Where(p => p.CanRead);
            var applicableProperties = readableProperties
                .Where(p => nameMatcher(p.Name))
                .Where(p =>
                    (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                    (p.GetCustomAttributes(typeof(IsDefault), true).Any()) ||
                    (!typeWasTranslatedFromVBScript && typeIsComVisible && !typeHasAmbiguousDispIdZeroMember && MemberHasDispIdZero(p))
                )
                .Select(p => p.GetGetMethod())
                .Where(m => matchesArgumentCount(m));

            // If no matches were found and we're looking for a parameter-less default member access and the target type was not translated from VBScript
            // source, then apply the other fallbacks that the C# compiler would have added to the IDispatch interface
            var allOptions = applicableMethods.Concat(applicableProperties);
            if (!allOptions.Any() && (defaultMemberBehaviour == DefaultMemberBehaviourOptions.MustBeDefault) && (numberOfArguments == 0) && !typeWasTranslatedFromVBScript && typeIsComVisible)
            {
                allOptions = allMethods.Where(m => !m.GetParameters().Any() && m.Name.Equals("Value", StringComparison.OrdinalIgnoreCase))
                    .Concat(readableProperties.Where(p => !p.GetIndexParameters().Any() && p.Name.Equals("Value", StringComparison.OrdinalIgnoreCase)).Select(p => p.GetGetMethod()));
                if (!allOptions.Any())
                    allOptions = new[] { type.GetMethod("ToString", Type.EmptyTypes) };
            }

            // In the cases where multiple options are identified, sort by the most specific (members declared on the current type those declared further
            // down in the inheritance tree)
            return allOptions.OrderByDescending(m => GetMemberInheritanceDepth(m, type));
        }

        private bool ParameterIsObjectParamsArray(ParameterInfo parameter)
        {
            if (parameter == null)
                throw new ArgumentNullException("parameter");

            return (parameter.ParameterType == typeof(object[])) && (parameter.GetCustomAttributes(typeof(ParamArrayAttribute), false).Length > 0);
        }

        private IEnumerable<MethodInfo> GetSetMethods(
            Type type,
            string optionalName,
            DefaultMemberBehaviourOptions defaultMemberBehaviour,
            MemberNameMatchBehaviourOptions memberNameMatchBehaviour,
            int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (!Enum.IsDefined(typeof(DefaultMemberBehaviourOptions), defaultMemberBehaviour))
                throw new ArgumentOutOfRangeException("defaultMemberBehaviour");
            if (!Enum.IsDefined(typeof(MemberNameMatchBehaviourOptions), memberNameMatchBehaviour))
                throw new ArgumentOutOfRangeException("memberNameMatchBehaviour");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            var nameMatcher = (optionalName != null) ? GetNameMatcher(optionalName, memberNameMatchBehaviour) : (name => true);
            return
                GetMethodsThatAreNotRelatedToProperties(type)
                    .Where(m => nameMatcher(m.Name))
                    .Where(m => m.GetParameters().Length == (numberOfArguments + 1)) // Method takes property arguments plus one for the value
                    .Where(m => m.GetCustomAttributes(true).Cast<Attribute>().Any(a => a is TranslatedProperty))
                    .Where(m =>
                        (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                        (m.GetCustomAttributes(true).Cast<Attribute>().Any(a => a is IsDefault))
                    )
                .Concat(
                    type.GetProperties()
                        .Where(p => p.CanWrite)
                        .Where(p => nameMatcher(p.Name))
                        .Where(p => p.GetIndexParameters().Length == numberOfArguments)
                        .Where(p =>
                            (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                            (p.GetCustomAttributes(true).Cast<Attribute>().Any(a => a is IsDefault))
                        )
                        .Select(p => p.GetSetMethod())
                );
        }

        private IEnumerable<MethodInfo> GetMethodsThatAreNotRelatedToProperties(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            return type.GetMethods() // This gets all public methods, whether they're declared in the specified type or anywhere in its inheritance tree
                .Except(type.GetProperties().Where(p => p.CanRead).Select(p => p.GetGetMethod()))
                .Except(type.GetProperties().Where(p => p.CanWrite).Select(p => p.GetSetMethod()));
        }

        private int GetMemberInheritanceDepth(MemberInfo memberInfo, Type type)
        {
            if (memberInfo == null)
                throw new ArgumentNullException("memberInfo");
            if (type == null)
                throw new ArgumentNullException("type");

            if (memberInfo.DeclaringType == type)
                return 0;
            if (type.BaseType == null)
                throw new ArgumentException("memberInfo is not declared on type or anywhere in its inheritance tree");
            return GetMemberInheritanceDepth(memberInfo, type.BaseType) + 1;
        }

        private Predicate<string> GetNameMatcher(string name, MemberNameMatchBehaviourOptions memberNameMatchBehaviour)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            if (!Enum.IsDefined(typeof(MemberNameMatchBehaviourOptions), memberNameMatchBehaviour))
                throw new ArgumentOutOfRangeException("memberNameMatchBehaviour");

            if (memberNameMatchBehaviour == MemberNameMatchBehaviourOptions.CaseInsensitive)
                return n => n.Equals(name, StringComparison.InvariantCultureIgnoreCase);
            if (memberNameMatchBehaviour == MemberNameMatchBehaviourOptions.Precise)
                return n => n.Equals(name, StringComparison.InvariantCultureIgnoreCase);
            if (memberNameMatchBehaviour == MemberNameMatchBehaviourOptions.UseNameRewriter)
            {
                var rewritterName = _nameRewriter(name);
                return n => n == rewritterName;
            }
            throw new ArgumentOutOfRangeException("memberNameMatchBehaviour");
        }

        private enum DefaultMemberBehaviourOptions
        {
            DoesNotMatter,
            MustBeDefault
        }

        private enum MemberNameMatchBehaviourOptions
        {
            CaseInsensitive,
            Precise,
            UseNameRewriter
        }

        private bool TypeIsTranslatedFromVBScript(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            return type.GetCustomAttributes(typeof(SourceClassName), true).Any();
        }

        private bool TypeIsComVisible(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            // For a type to be considered ComVisible, there must be a ComVisible(true) attribute on it or one of its explicitly derived types, having it
            // only on object (which everything implicitly derives from) is not sufficient since then everything would be ComVisible (since object has
            // the ComVisible(true) attribute set)
            if (type == typeof(object))
                return false;

            // We need to consider the ComVisible attributes on every class in the hierarchy (other than object - see above), so we specify inherit: false
            // here and walk up the tree until we find a ComVisible(true) - any ComVisible(false) settings encountered can be ignored so long as one of
            // the classes in the hierarchy has ComVisible(true).
            if (type.GetCustomAttributes(typeof(ComVisibleAttribute), inherit: false).Cast<ComVisibleAttribute>().Any(a => a.Value))
                return true;

            if (type.BaseType == null)
                return false;

            return TypeIsComVisible(type.BaseType);
        }

        private bool MemberHasDispIdZero(MemberInfo memberInfo)
        {
            if (memberInfo == null)
                throw new ArgumentNullException("memberInfo");

            return memberInfo.GetCustomAttributes(typeof(DispIdAttribute), inherit: true)
                .Cast<DispIdAttribute>()
                .Any(attribute => attribute.Value == 0);
        }

        private bool AnyDispIdZeroMemberExists(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            return type.GetMembers().Any(m => MemberHasDispIdZero(m));
        }

        private bool DispIdZeroIsAmbiguous(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            return type.GetMembers().Count(m => MemberHasDispIdZero(m)) > 1;
        }

        private sealed class InvokerCacheKey
        {
            private readonly int _hashCode;
            public InvokerCacheKey(object targetType, string optionalName, int numberOfArguments, bool onlyConsiderMethods)
            {
                if (targetType == null)
                    throw new ArgumentNullException("targetType");
                if (numberOfArguments < 0)
                    throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

                TargetType = targetType;
                OptionalName = optionalName;
                NumberOfArguments = numberOfArguments;
                OnlyConsiderMethods = onlyConsiderMethods;

                _hashCode = (TargetType.ToString() + "\n" + (optionalName ?? "") + "\n" + numberOfArguments.ToString() + "\n" + onlyConsiderMethods).GetHashCode();
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public object TargetType { get; private set; }

            /// <summary>
            /// This is optional and may be null
            /// </summary>
            public string OptionalName { get; private set; }

            /// <summary>
            /// This will always be zero or greater
            /// </summary>
            public int NumberOfArguments { get; private set; }

            /// <summary>
            /// Some member accesses will only consider methods - eg. "a.Name()" requires that "Name" be a method rather than a property on external references. Note
            /// that VBScript internally treats properties as method, so if "a" is a VBScript class with a public, readable, argument-less property "Name" then "a.Name()"
            /// will retrieve its value, as "a.Name" will. But if "a" is an IDispatch reference then it will only be queried for methods, so if it has a "Name" property
            /// then the call will fail if the interface does not also expose that property as a method.
            /// </summary>
            public bool OnlyConsiderMethods { get; private set; }

            public override int GetHashCode()
            {
                return _hashCode;
            }

            public override bool Equals(object obj)
            {
                if (obj == null)
                    throw new ArgumentNullException("obj");
                var cacheKey = obj as InvokerCacheKey;
                if (cacheKey == null)
                    return false;
                return (
                    (TargetType == cacheKey.TargetType) &&
                    (OptionalName == cacheKey.OptionalName) &&
                    (NumberOfArguments == cacheKey.NumberOfArguments)
                );
            }
        }

        private class ManagedEnumeratorWrapper : IEnumerable
        {
            private readonly IEnumerator _enumerator;
            public ManagedEnumeratorWrapper(IEnumerator enumerator)
            {
                if (enumerator == null)
                    throw new ArgumentNullException("enumerator");

                _enumerator = enumerator;
            }

            public IEnumerator GetEnumerator()
            {
                return _enumerator;
            }
        }

        private class IDispatchEnumeratorWrapper : IEnumerator
        {
            private readonly IEnumVariant _enumerator;
            private object _current;
            public IDispatchEnumeratorWrapper(IEnumVariant enumerator)
            {
                if (enumerator == null)
                    throw new ArgumentNullException("enumerator");

                _enumerator = enumerator;
                _current = null;
            }

            public object Current { get { return _current; } }

            public bool MoveNext()
            {
                uint fetched;
                _enumerator.Next(1, out _current, out fetched);
                return fetched != 0;
            }

            public void Reset()
            {
                _enumerator.Reset();
            }
        }

        [ComImport(), Guid("00020404-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumVariant
        {
            // From https://github.com/mosa/Mono-Class-Libraries/blob/master/mcs/class/CustomMarshalers/System.Runtime.InteropServices.CustomMarshalers/EnumeratorToEnumVariantMarshaler.cs
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Next(int celt, [MarshalAs(UnmanagedType.Struct)]out object rgvar, out uint pceltFetched);
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Skip(uint celt);
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Reset();
            [return: MarshalAs(UnmanagedType.Interface)]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            IEnumVariant Clone();
        }
    }
}
