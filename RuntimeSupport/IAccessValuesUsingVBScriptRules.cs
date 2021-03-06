﻿using System;
using System.Collections;
using System.Collections.Generic;

namespace VBScriptTranslator.RuntimeSupport
{
    public interface IAccessValuesUsingVBScriptRules
    {
        /// <summary>
        /// This is used to wrap arguments such that those that must be passed ByVal can have changes reflected after a method call completes
        /// </summary>
        IBuildCallArgumentProviders ARGS { get; }

        /// <summary>
        /// This requires a target with optional member accessors and arguments - eg. "Test" is a target only, "a.Test" has target "a" with one
        /// named member "Test", "a.Test(0)" has target "a", named member "Test" and a single argument "0". The expression "a(Test(0))" would
        /// require nested CALL executions, one with target "Test" and a single argument "0" and a second with target "a" and a single
        /// argument which was the result of the first call.
        /// </summary>
        object CALL(object context, object target, IEnumerable<string> members, IProvideCallArguments argumentProvider);
        
        /// <summary>
        /// This will throw an exception for null target or arguments references or if the setting fails (eg. invalid number of arguments,
        /// invalid member accessor - if specified - argument thrown by the target setter). This must not be called with a target reference
        /// only (null optionalMemberAccessor and zero arguments) as it would need to change the caller's reference to target, which is not
        /// possible (in that case, a straight assignment should be generated - no call to SET required). Note that the valueToSetTo argument
        /// comes before any others since VBScript will evaulate the right-hand side of the assignment before the left, which may be important
        /// if an error is raised at some point in the operation.
        /// </summary>
        void SET(object valueToSetTo, object context, object target, string optionalMemberAccessor, IProvideCallArguments argumentProvider);

        /// <summary>
        /// This will never throw an exception, a value is either considered by VBScript to be a value type (including values such as Empty,
        /// Null, numbers, dates, strings, arrays) or not
        /// </summary>
        bool IsVBScriptValueType(object o);

        /// <summary>
        /// Try to reduce a reference down to a value type, applying VBScript defaults logic - returns true if this is possible, returns false if it
        /// is not possible and throws an exception if one was raised while a default member was being evaluated (where applicable). Ift rue is
        /// returned then the out asValueType argument will be the manipulated value. If false is returned but the target reference had a default
        /// parameterless member, then the out asValueType argument will be given the value of this member (even though it was determined that
        /// this reference was not a value type. If false is returned and the target reference did not have a parameterless default member,
        /// then the out asValueType argument will be a reference to the target (set to the same reference that the "o" argument has).
        /// </summary>
        bool TryVAL(object o, out bool parameterLessDefaultMemberWasAvailable, out object asValueType);

        /// <summary>
        /// Reduce a reference down to a value type, applying VBScript defaults logic - thrown an exception if this is not possible (null is
        /// acceptable as an input and corresponding return value)
        /// </summary>
        object VAL(object o, string exceptionMessageForInvalidContent = null);

        /// <summary>
        /// This will only return a non-VBScript-value-type, if unable to then an exception will be raised (this is used to wrap the right-hand
        /// side of a SET assignment)
        /// </summary>
        object OBJ(object o, string optionalExceptionMessageForInvalidContent = null);

        /// <summary>
        /// Reduce a reference down to a boolean value type, in the same ways that VBScript would attempt.
        /// </summary>
        bool BOOL(object o, string optionalExceptionMessageForInvalidContent = null);

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
        object NUM(object o, params object[] numericValuesTheTypeMustBeAbleToContain);

        /// <summary>
        /// This wraps a call to NUM and allows an exception to be made for DBNull.Value (VBScript Null) in that the same value will be returned
        /// (it is not a valid input for NUM).
        /// </summary>
        object NullableNUM(object o);

        /// <summary>
        /// This is similar to NullableNUM in that it is used for comparisons involving date literals, where the other side has to be interpreted as
        /// a date but must also support null. It wraps DATE (and so supports all VBScript date parsing methods).
        /// </summary>
        object NullableDATE(object o);

        /// <summary>
        /// Reduce a reference down to a date value type, applying VBScript defaults logic and then taking a date representation. Numeric values are
        /// acceptable (taken as the number of days since the zero date, with support for fractional values). String values are acceptable, they will be
        /// parsed using VBScript's rules (and taking into account culture, where applicable). Null is acceptable and will return in the zero date being
        /// returned, DBNull.Value (aka VBScript Null) is not acceptable and will return in an exception being raised.
        /// </summary>
        DateTime DATE(object o, string optionalExceptionMessageForInvalidContent = null);

        /// <summary>
        /// Apply the same logic as STR but allow DBNull.Value (returning it back). This conversion should only used for comparisons with string literals,
        /// where special rules apply (which makes the method slightly less useful than NUM, which is used in comparisons with numeric literals but also
        /// in some other cases, such as FOR loops).
        /// </summary>
        object NullableSTR(object o);

        /// <summary>
        /// Reduce a reference down to a string value type, applying VBScript defaults logic and then taking a string representation. Null is acceptable
        /// and will return in a blank string being returned, DBNull.Value (aka VBScript Null) is not acceptable and will return in an exception being
        /// raised.
        /// </summary>
        string STR(object o, string optionalExceptionMessageForInvalidContent = null);

        /// <summary>
        /// Reduce a reference down to a boolean, throwing an exception if this is not possible. This will apply the same logic as VAL but then
        /// require a numeric value or null, otherwise an exception will be raised. Zero and null equate to false, non-zero numbers to true.
        /// </summary>
        bool IF(object o);

        /// <summary>
        /// Layer an enumerable wrapper over a reference, if possible (an exception will be thrown if not)
        /// </summary>
        IEnumerable ENUMERABLE(object o);
    }
}
