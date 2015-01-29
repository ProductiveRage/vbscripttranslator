using CSharpSupport.Implementations;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace CSharpSupport
{
    public class DefaultRuntimeSupportClassFactory
    {
        private static Regex _multipleUnderscoreCondenser;
        static DefaultRuntimeSupportClassFactory()
        {
            _multipleUnderscoreCondenser = new Regex("_{2,}", RegexOptions.Compiled);
            DefaultNameRewriter = RewriteName;
            DefaultVBScriptValueRetriever = new VBScriptEsqueValueRetriever(DefaultNameRewriter);
        }

        /// <summary>
        /// Each compat functionality provider instance should be disposed of after the request has completed to ensure that any managed resources are tidied
        /// up (this is an approximation of VBScript's deterministic reference-counting garbage collector - it doesn't dispose of the resources as quickly,
        /// but at least they're guaranteed to be dealt with after the request ends if this is disposed).
        /// </summary>
        public static IProvideVBScriptCompatFunctionalityToIndividualRequests Get()
        {
            return new DefaultRuntimeFunctionalityProvider(DefaultNameRewriter, DefaultVBScriptValueRetriever);
        }

        /// <summary>
        /// This is a static reference as it will build up an member access cache so that subsequent requests for a given method signature on a particular
        /// type does not need to be determined from scratch. The implementation is thread safe and may be shared between requests.
        /// </summary>
        public static IAccessValuesUsingVBScriptRules DefaultVBScriptValueRetriever { get; private set; }

        /// <summary>
        /// This is a static reference since its implementation does not change and may be shared across all requests. It has no state and so is thread safe.
        /// </summary>
        public static Func<string, string> DefaultNameRewriter { get; private set; }

        private static string RewriteName(string value)
        {
            if (value == null)
                throw new ArgumentNullException("value");

            // If the value is a valid C# name then do nothing
            // - The specification for valid identifiers may be found at https://msdn.microsoft.com/en-us/library/aa664670(VS.71).aspx, but I'm going
            //   for an extremely simple approach here that may incorrectly identify some valid names as invalid (and so manipulate them further down)
            //   but it should do the job well enough (it can always be revisited if problems are found with particular data)
            // Note: VBScript allows all sorts of surprising names if they're wrapped in square brackets. The square brackets do not become part of
            // the name, so "Dim [i]" is exactly the same as "Dim i", but it does allow for otherwise-unsupported characters to be used (such as
            // white space). In fact, with square brackets, a blank name is acceptable! (Hence the Length check below).
            if ((value.Length > 0)
            && ((value[0] == '_') || char.IsLetter(value[0]))
            && value.All(c => (c == '_') || char.IsLetterOrDigit(c)))
                return value;

            // If we have to manipulate the value then we'll perform some replacements and then append the hash code from the original string to the
            // end of it to try to avoid collisions. This will not be perfect but it should be good enough. (Probably the only fool-proof approach
            // would be for a name rewriter to be used during the translation process that guarantees that it produces unique names and then for
            // the mappings to be exported and incorporated into a name rewriter for use when the translate code is executed - this, though, would
            // be complicated since there would probably need to be additional data about the context where the name translation takes places as
            // well as the actual source name; this is why I'm happy to stick with this simpler approach for now).
            if (value == "")
            {
                // Special case for blank string since GetHash returns zero - I'll make up a psuedo-random name that hopefully won't clash with
                // anything else (not very scientific, I know!) http://xkcd.com/221/
                return "unnamedVariable027396729921";
            }
            var rewrittenValue = new string(
                value.Select(c => char.IsLetterOrDigit(c) ? c : '_')
                .ToArray()
            );
            rewrittenValue = _multipleUnderscoreCondenser.Replace(rewrittenValue, "_").Trim('_');
            if (!char.IsLetter(rewrittenValue[0]))
                rewrittenValue = "x" + rewrittenValue;
            return rewrittenValue += GetHash(value);
        }

        /// <summary>
        /// I could have use string.GetHashCode but that is allowed to vary between version of .net - so translated code emitted in one version may
        /// fail when running under another version when names have to be rewritten. This is probably quite unlikely but I thought taking a bog
        /// standard alternative implementation (this is Jenkins one-at-a-time, see http://en.wikipedia.org/wiki/Jenkins_hash_function) would make
        /// for a reasonable starting point for a consistent-across-framework-versions approach.
        /// </summary>
        private static uint GetHash(string value)
        {
            uint hash = 0;
            for (var i = 0; i < value.Length; ++i)
            {
                hash += Convert.ToUInt32(value[i]);
                hash += (hash << 10);
                hash ^= (hash >> 6);
            }
            hash += (hash << 3);
            hash ^= (hash >> 11);
            hash += (hash << 15);
            return hash;
        }
    }
}
