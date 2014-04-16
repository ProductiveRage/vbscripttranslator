using System;
using System.Collections.Generic;

namespace CSharpSupport
{
	public static class IBuildCallArgumentProviders_Extensions
    {
        /// <summary>
        /// TODO
        /// This should return a reference to itself to enable chaining when building up argument sets
        /// </summary>
		public static IBuildCallArgumentProviders RefIfArray(this IBuildCallArgumentProviders source, object target, params IProvideCallArguments[] argumentProviders)
		{
			if (source == null)
				throw new ArgumentNullException("source");
			if (target == null)
				throw new ArgumentNullException("target");
			if (argumentProviders == null)
				throw new ArgumentNullException("argumentProviders");

			return source.RefIfArray(target, (IEnumerable<IProvideCallArguments>)argumentProviders);
		}
    }
}
