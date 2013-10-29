using System;

namespace CSharpWriter.Logging
{
    public class ConsoleLogger : ILogInformation
    {
        public void Warning(string content)
        {
            Console.WriteLine(content);
        }
    }
}
