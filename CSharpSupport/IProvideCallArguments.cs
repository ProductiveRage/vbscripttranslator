using System.Collections.Generic;

namespace CSharpSupport
{
    /*
        Definitely ByVal
        - Constant (numeric, string or builtin - eg. Nothing)
          > BuiltInValueExpressionSegment
          > NumericValueExpressionSegment
          > StringValueExpressionSegment
        - Bracketed section
          > BracketedExpressionSegment
        - Confirmed function call (token matched to known function)
        - Confirmed property accessor (explicit property accessor in tokens)
        - New instance
          > NewInstanceExpressionSegment

        Definitely ByRef *
        - Non-bracketed variable access without brackets or properties (token matched to known variable)

        Possibly ByRef **
        - Non-bracketed variable access with brackets
          > eg. "a(0)
          > eg. "a(0, 0)
          > eg. "a(0)(0)"
          if all accesses are array elements, none are default property or method accesses

        * This is the only one that requires a setter delegate

        ** Should element index accessors be cached to guarantee consistency?
           - Suggests that this testing be done before the function call
             > Yes, record in initial values set along with updater delegate to be consistent with Definitely-ByRef entries?
     */
    public interface IProvideCallArguments
    {
        /// <summary>
        /// This will always be zero or greater
        /// </summary>
        int NumberOfArguments { get; }

        /// <summary>
        /// The presence of brackets in the source following a member access with zero arguments may affect the available calling mechanisms on
        /// the target, so this is important information to record and expose
        /// </summary>
        bool UseBracketsWhereZeroArguments { get; }

        /// <summary>
        /// This will always return a set with NumberOfArguments items in it
        /// </summary>
        IEnumerable<object> GetInitialValues();

        /// <summary>
        /// The index must be zero or greater and less than NumberOfArguments. If the argument at that index may not be overrwritten then the
        /// function call will have no effect.
        /// </summary>
        void OverwriteValueIfByRef(int index, object value);
    }
}
