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

            // C# already uses Double with decimal numbers, so we don't need any special case if the number is expressed as a decimal with numbers both before
            // and after the decimal point (eg. "1.2"). However, VBScript throws another curve ball and supports numbers with a decimal point with no digits
            // after it (eg. "1."). This is not valid in C# ("Identifier expected") so we have to slap a zero on the end (making it "1.0", which will be
            // defined as a double). Note that there is no such issue when leading with the decimal point (".1" is valid VBScript AND C# code).
            if (token.Content.Contains("."))
                return token.Content + (token.Content.EndsWith(".") ? "0" : "");

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

        /// <summary>
        /// During the translation process, it may be necessary to wrap up numeric literals such that they are not then considered to be numeric literals
        /// when further processing is done (see the OperatorCombiner and the StatementTranslator for details about cases where this may happen - around
        /// comparisons of values / literals). The numeric value dictates what wrapper function will be appropriate and not affect its type - if a VBScript
        /// "Integer" is wrapped in a CDbl call, then the return type is changed to "Double". If VBScript considers a value to be an "Integer" then it must
        /// be wrapped in a CInt call in order to prevent changing the value's type. This function determines what built-in function is acceptable. It will
        /// never return null and it will only throw an exception if the token argument is null since all numbers can be safely wrapped.
        /// </summary>
        public static string GetSafeWrapperFunctionName(this NumericValueToken token)
        {
            if (token == null)
                throw new ArgumentNullException("token");

            // If it's a decimal then we need to use CDbl
            if (token.Content.Contains("."))
                return "CDbl";
            if ((token.Value >= Int16.MinValue) && (token.Value <= Int16.MaxValue))
                return "CInt";
            if ((token.Value >= Int32.MinValue) || (token.Value <= Int32.MaxValue))
                return "CLng";
            return "CDbl";
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
