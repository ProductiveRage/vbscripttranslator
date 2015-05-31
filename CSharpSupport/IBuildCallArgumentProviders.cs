using System;
using System.Collections.Generic;

namespace CSharpSupport
{
    public interface IBuildCallArgumentProviders
    {
        /// <summary>
        /// Add an argument to the set that will be passed by-val, regardless of whether the target call site expects a by-ref or by-val argument.
        /// This should return a reference to itself to enable chaining when building up argument sets.
        /// </summary>
        IBuildCallArgumentProviders Val(object value);

        /// <summary>
        /// Add an argument to the set that will be passed by-ref if the target call site expects a by-ref argument (otherwise it will be treated as
        /// by-val). This should return a reference to itself to enable chaining when building up argument sets.
        /// </summary>
        IBuildCallArgumentProviders Ref(object value, Action<object> valueUpdater);

        /// <summary>
        /// Add an argument to the set that will be passed by-ref if the target call site expects a by-ref argument and if the value is an array
        /// (otherwise it will be treated as by-val). This should return a reference to itself to enable chaining when building up argument sets.
        /// </summary>
		IBuildCallArgumentProviders RefIfArray(object target, IEnumerable<IProvideCallArguments> argumentProviders);
        
        /// <summary>
        /// Specify that brackets were specified, even if there were zero arguments - this may affect the available call mechanisms (eg. it will
        /// only access methods, not properties, on IDispatch targets). This information is only of use if there are zero arguments (though it
        /// is not invalid to indicate this even where there are arguments present). This should return a reference to itself to enable chaining
        /// when building up argument sets.
        /// </summary>
        IBuildCallArgumentProviders ForceBrackets();

        /// <summary>
        /// This will never return null
        /// </summary>
        IProvideCallArguments GetArgs();
    }
}
