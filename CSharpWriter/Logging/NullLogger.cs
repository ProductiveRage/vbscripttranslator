﻿namespace VBScriptTranslator.CSharpWriter.Logging
{
    public class NullLogger : ILogInformation
    {
        public void Warning(string content) { }
    }
}
