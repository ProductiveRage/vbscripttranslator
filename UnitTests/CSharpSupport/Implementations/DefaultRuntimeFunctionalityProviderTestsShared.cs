using CSharpSupport.Implementations;
using System;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public partial class DefaultRuntimeFunctionalityProviderTests
    {
        private static DefaultRuntimeFunctionalityProvider GetDefaultRuntimeFunctionalityProvider()
        {
            Func<string, string> nameRewriter = name => name;
            return new DefaultRuntimeFunctionalityProvider(
                nameRewriter,
                new VBScriptEsqueValueRetriever(nameRewriter)
            );
        }
    }
}
