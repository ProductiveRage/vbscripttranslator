using System.Collections.Generic;

namespace CSharpSupport
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
        object CALL(object target, IEnumerable<string> members, IProvideCallArguments argumentProvider);
        
        /// <summary>
        /// This will throw an exception for null target or arguments references or if the setting fails (eg. invalid number of arguments,
        /// invalid member accessor - if specified - argument thrown by the target setter). This must not be called with a target reference
        /// only (null optionalMemberAccessor and zero arguments) as it would need to change the caller's reference to target, which is not
        /// possible (in that case, a straight assignment should be generated - no call to SET required). Note that the valueToSetTo argument
        /// comes before any others since VBScript will evaulate the right-hand side of the assignment before the left, which may be important
        /// if an error is raised at some point in the operation.
        /// </summary>
        void SET(object valueToSetTo, object target, string optionalMemberAccessor, IProvideCallArguments argumentProvider);

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
        /// Reduce a reference down to a numeric value type, applying VBScript defaults logic and then trying to parse as a number - throwing
        /// an exception if this is not possible. Null (aka VBScript Empty) is acceptable and will result in zero being returned. DBNull.Value
        /// (aka VBScript Null) is not acceptable and will result in an exception being raised, as any other invalid value (eg. a string or
        /// an object without an appropriate default property member) will.
        /// </summary>
        double NUM(object o);
    
        /// <summary>
        /// Reduce a reference down to a boolean, throwing an exception if this is not possible. This will apply the same logic as VAL but then
        /// require a numeric value or null, otherwise an exception will be raised. Zero and null equate to false, non-zero numbers to true.
        /// </summary>
        bool IF(object o);
    }
}
