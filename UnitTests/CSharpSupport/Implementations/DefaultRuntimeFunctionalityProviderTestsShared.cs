using CSharpSupport.Implementations;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public partial class DefaultRuntimeFunctionalityProviderTests
    {
        private static DefaultRuntimeFunctionalityProvider GetDefaultRuntimeFunctionalityProvider()
        {
            return new DefaultRuntimeFunctionalityProvider(name => name);
        }
    }
}
