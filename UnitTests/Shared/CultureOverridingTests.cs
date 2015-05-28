using System;
using System.Globalization;
using System.Threading;

namespace VBScriptTranslator.UnitTests.Shared
{
    public abstract class CultureOverridingTests
    {
        private readonly CultureInfo _originalCulture;
        protected CultureOverridingTests(CultureInfo culture)
        {
            if (culture == null)
                throw new ArgumentNullException("culture");

            _originalCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = culture;
        }
        ~CultureOverridingTests()
        {
            Thread.CurrentThread.CurrentCulture = _originalCulture;
        }
    }
}
