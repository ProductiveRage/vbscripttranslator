using System;
using System.Collections.Generic;

namespace CSharpSupport
{
    public interface IBuildCallArgumentProviders
    {
        /// <summary>
        /// TODO
        /// This should return a reference to itself to enable chaining when building up argument sets
        /// </summary>
        IBuildCallArgumentProviders Val(object value);

        /// <summary>
        /// TODO
        /// This should return a reference to itself to enable chaining when building up argument sets
        /// </summary>
        IBuildCallArgumentProviders Ref(object value, Action<object> valueUpdater);

        /// <summary>
        /// TODO
        /// This should return a reference to itself to enable chaining when building up argument sets
        /// </summary>
		IBuildCallArgumentProviders RefIfArray(object target, IEnumerable<IProvideCallArguments> argumentProviders);

        /// <summary>
        /// TODO
        /// </summary>
        IProvideCallArguments GetArgs();
    }
}
