using System;

namespace Tester
{
    /// <summary>
    /// Most of the early test scripts that I'm writing assume a WScript reference, and an Echo method on that reference. This will work as that reference
    /// (to be passed in as the EnvironmentReferences for the translated programs).
    /// </summary>
    public class WScriptMock
    {
        public void Echo(object content)
        {
            Console.WriteLine(content);
        }
    }
}
