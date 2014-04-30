using System.Collections.Generic;

namespace CSharpSupport
{
    public interface IAccessValuesUsingVBScriptRules
    {
        /// <summary>
        /// TODO
        /// </summary>
        IBuildCallArgumentProviders ARGS { get; }

        /// <summary>
        /// This requires a target with optional member accessors and arguments - eg. "Test" is a target only, "a.Test" has target "a" with one
        /// named member "Test", "a.Test(0)" has target "a", named member "Test" and a single argument "0". The expression "a(Test(0))" would
        /// require nested CALL executions, one with target "Test" and a single argument "0" and a second with target "a" and a single
        /// argument which was the result of the first call.
        /// </summary>
        object CALL(object target, IEnumerable<string> members, IProvideCallArguments argumentProvider);
        
        /// <summary>
        /// This will throw an exception for null target or arguments references or if the setting fails (eg. invalid number of arguments,
        /// invalid member accessor - if specified - argument thrown by the target setter). This must not be called with a target reference
        /// only (null optionalMemberAccessor and zero arguments) as it would need to change the caller's reference to target, which is not
        /// possible (in that case, a straight assignment should be generated - no call to SET required).
        /// </summary>
        void SET(object target, string optionalMemberAccessor, IProvideCallArguments argumentProvider, object value);

        /// <summary>
        /// Reduce a reference down to a value type, applying VBScript defaults logic - thrown an exception if this is not possible (null is
        /// acceptable as an input and corresponding return value)
        /// </summary>
        object VAL(object o);

        /// <summary>
        /// This will only return a non-VBScript-value-type, if unable to then an exception will be raised (this is used to wrap the right-hand
        /// side of a SET assignment)
        /// </summary>
        object OBJ(object o);

        /// <summary>
        /// Reduce a reference down to a boolean, throwing an exception if this is not possible. This will apply the same logic as VAL but then
        /// require a numeric value or null, otherwise an exception will be raised. Zero and null equate to false, non-zero numbers to true.
        /// </summary>
        bool IF(object o);
    }
}
