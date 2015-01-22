using System;

namespace VBScriptTranslator.LegacyParser.Tokens.Basic
{
    public static class NumericValueToken_Extensions
    {
        /// <summary>
        /// When dealing with number literals, VBScript determines what type to use based upon whether it's a decimal or integer and then the size of the integer.
        /// If it's a decimal it's always a "Double" (in VBScript parlance). If it's an integer then it will be a "Integer" if it's small enough (Int16) or "Long"
        /// if it's bigger but not much (Int32 / int) and after that it goes to being a Double. This method returns C# code that describes the value and ensures
        /// that it is cast to the appropriate type (in order to be consistent with VBScript).
        /// </summary>
        public static string AsCSharpValue(this NumericValueToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            //var numericVlau
            //var aa = 1.;
            // C# already uses Double with decimal numbers, so we don't need any special case there
            if (token.Content.Contains("."))
                return token.Content;

            // C# will default to int (Int32) for integers, we need to override this for smaller values
            if ((token.Value >= Int16.MinValue) && (token.Value <= Int16.MaxValue))
                return "(Int16)" + token.Content;
            
            // When Int32 would overflow, C# will bump to Int64, we need to override this to use Double.
            if ((token.Value < Int32.MinValue) || (token.Value > Int32.MaxValue))
                return token.Content + "d";

            // The only other case is when it's an integer in the range (between Int16 and Int32) where VBScript would jump to a
            // "Long", which is "int" in .net - so no funny business required
            return token.Content;
        }

        /// <summary>
        /// Some numbers can be "unwrapped" from VBScript number-casting functions if it would not alter the number's type - eg. in VBScript, the number 1 is
        /// considered an Integer, which is what CInt(1) returns, so CInt(1) can be replaced with the number 1 with no functional effect. However, CLng(1)
        /// could NOT be replaced with 1 since the type would become "Integer" instead of "Long". This method reports whether or not a particular unwrapping
        /// is acceptable for a given NumericValueToken or not.
        /// </summary>
        public static bool IsSafeToUnwrapFrom(this NumericValueToken token, BuiltInFunctionToken function)
        {
            if (token == null)
                throw new ArgumentNullException("token");
            if (function == null)
                throw new ArgumentNullException("function");

            // This basically takes the rules from AsCSharpValue() and inverts them
            if (token.Content.Contains(".") && function.Content.Equals("CDbl", StringComparison.OrdinalIgnoreCase))
                return true;
            if ((token.Value >= Int16.MinValue) && (token.Value <= Int16.MaxValue) && function.Content.Equals("CInt", StringComparison.OrdinalIgnoreCase))
                return true;
            if ((token.Value >= Int32.MinValue) || (token.Value <= Int32.MaxValue))
            {
                if (function.Content.Equals("CLng", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return function.Content.Equals("CDbl", StringComparison.OrdinalIgnoreCase);
        }

        public static BuiltInFunctionToken GetSafeWrapperFunction(this NumericValueToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            throw new NotImplementedException(); // TODO
        }

        public static NumericValueToken GetNegative(this NumericValueToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            if (token.Content.StartsWith("-"))
                return new NumericValueToken(token.Content.Substring(1), token.LineIndex);
            return new NumericValueToken("-" + token.Content, token.LineIndex);
        }
    }
}
