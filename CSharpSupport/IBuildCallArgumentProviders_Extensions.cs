using System;
using System.Linq;

namespace VBScriptTranslator.RuntimeSupport
{
	public static class IBuildCallArgumentProviders_Extensions
    {
        /// <summary>
        /// TODO
        /// This should return a reference to itself to enable chaining when building up argument sets
        /// </summary>
        public static IBuildCallArgumentProviders RefIfArray(this IBuildCallArgumentProviders source, object target, params IBuildCallArgumentProviders[] argumentProviderBuilders)
        {
            if (source == null)
                throw new ArgumentNullException("source");
            if (target == null)
                throw new ArgumentNullException("target");
            if (argumentProviderBuilders == null)
                throw new ArgumentNullException("argumentProviders");

            var argumentProviders = argumentProviderBuilders.Select(b => (b == null) ? null : b.GetArgs()).ToArray();
            if (argumentProviders.Any(p => p == null))
                throw new ArgumentException("Null reference encountered in argumentProviderBuilders set");

            return source.RefIfArray(target, argumentProviders);
        }
    }
}
