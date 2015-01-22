using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Exceptions;
using CSharpSupport.Implementations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace CSharpSupport.Implementations
{
    /// <summary>
    /// This is intended to be built up over time and used with classes that are output by the translator. Clearly, at this point, it is noticeably lacking
    /// in working functionality, but it provides something, at least, to use to test the very simple programs that can be translated at this time.
    /// </summary>
    public class DefaultRuntimeFunctionalityProvider : VBScriptEsqueValueRetriever, IProvideVBScriptCompatFunctionalityToIndividualRequests
    {
        public DefaultRuntimeFunctionalityProvider(Func<string, string> nameRewriter) : base(nameRewriter) { }

        // Arithemetic operators
        public double POW(object l, object r) { throw new NotImplementedException(); }
        public double DIV(object l, object r) { throw new NotImplementedException(); }
        public double MULT(object l, object r) { throw new NotImplementedException(); }
        public int INTDIV(object l, object r) { throw new NotImplementedException(); }
        public double MOD(object l, object r) { throw new NotImplementedException(); }
        public double ADD(object l, object r) { throw new NotImplementedException(); }
        public double SUBT(object o) { throw new NotImplementedException(); }
        public double SUBT(object l, object r) { throw new NotImplementedException(); }

        // String concatenation
        public string CONCAT(object l, object r) { throw new NotImplementedException(); }

        // Logical operators (these return VBScript Null if one or both sides of the comparison are VBScript Null)
        public object NOT(object o) { throw new NotImplementedException(); }
        public object AND(object l, object r) { throw new NotImplementedException(); }
        public object OR(object l, object r) { throw new NotImplementedException(); }
        public object XOR(object l, object r) { throw new NotImplementedException(); }

        // Comparison operators (these return VBScript Null if one or both sides of the comparison are VBScript Null)
        /// <summary>
        /// This will return DBNull.Value or boolean value. VBScript has rules about comparisons between "hard-typed" values (aka literals), such
        /// that a comparison between (a = 1) requires that the value "a" be parsed into a numeric value (resulting in a Type Mismatch if this is
        /// not possible). However, this logic must be handled by the translation process before the EQ method is called. Both comparison values
        /// must be treated as non-object-references, so if they are not when passed in then the method will try to retrieve non-object values
        /// from them - if this fails then a Type Mismatch error will be raised. If there are no issues in preparing both comparison values,
        /// this will return DBNull.Value if either value is DBNull.Value and a boolean otherwise.
        /// </summary>
        public object EQ(object l, object r) { return ToVBScriptNullable(EQ_Internal(l, r)); }
        private bool? EQ_Internal(object l, object r)
        {
            // Both sides of the comparison must be simple VBScript values (ie. not object references) - pushing both values through VAL will handle
            // that (an exception will be raised if this operation fails and the value will not be affect if it was already an acceptable type)
            l = VAL(l);
            r = VAL(r);
            
            // Let's get the outliers out of the way; VBScript Null and Empty..
            if ((l == DBNull.Value) || (r == DBNull.Value))
                return null; // If one or both sides of the comparison are "Null" then this is what is returned
            if ((l == null) && (r == null))
                return true; // If both sides are Empty then they are considered to match
            else if ((l == null) || (r == null))
            {
                // The default values of VBScript primitives (number, strings and booleans) are considered to match Empty
                var nonNullValue = l ?? r;
                if ((base.IsDotNetNumericType(nonNullValue) && (Convert.ToDouble(nonNullValue)) == 0)
                || ((nonNullValue as string) == "")
                || ((nonNullValue is bool) && !(bool)nonNullValue))
                    return true;
                return false;
            }

            // Booleans have some funny behaviour in that they will match values of other types (numbers, but not strings unless string literals
            // are in the comparison, which is not logic that this method has to deal with). If one of the values is a boolean and the other isn't,
            // and none of the special cases are met, then there must not be a match.
            if ((l is bool) && (r is bool))
                return (bool)l == (bool)r;
            else if ((l is bool) || (r is bool))
            {
                var boolValue = (bool)((l is bool) ? l : r);
                var nonBoolValue = (l is bool) ? r : l;
                if (!base.IsDotNetNumericType(nonBoolValue))
                    return false;
                return (boolValue && (Convert.ToDouble(nonBoolValue) == -1)) || (!boolValue && (Convert.ToDouble(nonBoolValue) == 0));
            }

            // Now consider numbers on one or both sides - all special cases are out of the way now so they're either equal or they're not (both
            // sides must be numbers, otherwise it's a non-match)
            if (base.IsDotNetNumericType(l) && base.IsDotNetNumericType(r))
                return Convert.ToDouble(l) == Convert.ToDouble(r);
            else if (base.IsDotNetNumericType(l) || base.IsDotNetNumericType(r))
                return false;

            // Now do the same for strings and then dates - same deal; they must have consistent types AND values
            if ((l is string) && (r is string))
                return (string)l == (string)r;
            else if ((l is string) || (r is string))
                return false;
            if ((l is DateTime) && (r is DateTime))
                return (DateTime)l == (DateTime)r;

            // Frankly, if we get here then I have no idea what's happened. It will be much easier to identify issues (if any are encountered) if an
            // exception is raised rather than a false response return
            throw new NotSupportedException("Don't know how to compare values of type " + TYPENAME(l) + " and " + TYPENAME(r));
        }

        public object NOTEQ(object l, object r)
        {
            // We can just reverse EQ_Internal's result here, unless it returns null - if it returns null then it means that comparison was not
            // meaningful (one or both sides were DBNull.Value) and so DBNull.Value should be returned.
            var opposingEqualityResult = EQ_Internal(l, r);
            if (opposingEqualityResult == null)
                return null;
            return !opposingEqualityResult.Value;
        }

        public object LT(object l, object r) { return ToVBScriptNullable(LT_Internal(l, r, allowEquals: false)); }
        public object LTE(object l, object r) { return ToVBScriptNullable(LT_Internal(l, r, allowEquals: true)); }
        private bool? LT_Internal(object l, object r, bool allowEquals)
        {
            // Both sides of the comparison must be simple VBScript values (ie. not object references) - pushing both values through VAL will handle
            // that (an exception will be raised if this operation fails and the value will not be affect if it was already an acceptable type)
            l = VAL(l);
            r = VAL(r);

            // If one or both sides of the comparison as VBScript Null then that is what is returned
            if ((l == DBNull.Value) || (r == DBNull.Value))
                return null;

            // Check the equality case first, since there may be an early exit we can make (this should return a true or false since the "Null" cases
            // have been handled) - if the values ARE equal then either return true (if allowEquals is true) or false (if allowEquals is false). If
            // not then we'll have to do more work.
            var eq = EQ_Internal(l, r);
            if (eq == null)
                throw new NotSupportedException("Don't know how to compare values of type " + TYPENAME(l) + " and " + TYPENAME(r));
            if (eq.Value)
                return allowEquals;

            // Deal with string special cases next - if both are strings then perform a string comparison. If only one is a string, and it is not blank,
            // then that value is bigger (so if it's on the left then return false and if it's on the right then return true). Blank strings get special
            // handling and are effectively treated as zero (see further down).
            var lString = l as string;
            var rString = r as string;
            if ((lString != null) && (rString != null))
            {
                var stringComparisonResult = STRCOMP_Internal(lString, rString, 0);
                if ((stringComparisonResult == null) || (stringComparisonResult.Value == 0))
                    throw new NotSupportedException("Don't know how to compare values of type " + TYPENAME(l) + " and " + TYPENAME(r));
                return stringComparisonResult.Value < 0;
            }
            if ((lString != null) && (lString != ""))
                return false;
            if ((rString != null) && (rString != ""))
                return true;

            // Now we should only have values which can treated as numeric
            // - Actual numbers
            // - Booleans (which return zero or minus one when passed through CDBL)
            // - Null aka VBScript Empty (which returns zero when passed through CDBL)
            // - Blank strings (which can not be passed through CDBL without causing an error, but which we can treat as zero)
            var lNumeric = (lString == "") ? 0 : CDBL(l);
            var rNumeric = (rString == "") ? 0 : CDBL(r);
            return lNumeric < rNumeric;
        }

        public object GT(object l, object r) { return ToVBScriptNullable(GT_Internal(l, r, allowEquals: false)); }
        public object GTE(object l, object r) { return ToVBScriptNullable(GT_Internal(l, r, allowEquals: true)); }
        private bool? GT_Internal(object l, object r, bool allowEquals)
        {
            // This can just LT_Internal, rather than trying to deal with too much logic itself. When calling LT_Internal, the "allowEquals" value must be
            // the opposite of what we have here - if we are considering GTE then we want !LT (since the equality case should be a match here and not a
            // result which is inverted), if we are considering GT here then we want !LTE (since then equality case would not be a match and LTE would
            // return true for equal l and r values and we would want to invert that result). If LT_Internal returns null, then it means that the
            // comparison is not meaningful (in other words, DBNull.Value was on one or both sides and so DBNull.Value should be returned for
            // any comparison - whether EQ, NOTEQ, LT, GT, etc..)
            var opposingLessThanResult = LT_Internal(l, r, !allowEquals);
            if (opposingLessThanResult == null)
                return null;
            return !opposingLessThanResult.Value;
        }

        public object IS(object l, object r) { throw new NotImplementedException(); }
        public object EQV(object l, object r) { throw new NotImplementedException(); }
        public object IMP(object l, object r) { throw new NotImplementedException(); }

        // Builtin functions - TODO: These are not fully specified yet (eg. LEFT requires more than one parameter and INSTR requires multiple parameters and
        // overloads to deal with optional parameters)
        // - Type conversions
        public byte CBYTE(object value) { return GetAsNumber<byte>(value, Convert.ToByte); }
        public object CBOOL(object value) { throw new NotImplementedException(); }
        public double CDBL(object value) { return GetAsNumber<double>(value, Convert.ToDouble); }
        public object CDATE(object value) { throw new NotImplementedException(); }
        public Int16 CINT(object value) { return GetAsNumber<Int16>(value, Convert.ToInt16); } // TODO: Confirm appropriate response type (Int16?)
        public int CLNG(object value) { return GetAsNumber<int>(value, Convert.ToInt32); }
        public float CSNG(object value) { return GetAsNumber<float>(value, Convert.ToSingle); }
        public string CSTR(object value) { throw new NotImplementedException(); }
        public string INT(object value) { throw new NotImplementedException(); }
        public string STRING(object value) { throw new NotImplementedException(); }
        // - String functions
        public object ASC(object value) { throw new NotImplementedException(); }
        public object ASCB(object value) { throw new NotImplementedException(); }
        public object ASCW(object value) { throw new NotImplementedException(); }
        public object CHR(object value) { throw new NotImplementedException(); }
        public object CHRB(object value) { throw new NotImplementedException(); }
        public object CHRW(object value) { throw new NotImplementedException(); }
        public object INSTR(object value) { throw new NotImplementedException(); }
        public object INSTRREV(object value) { throw new NotImplementedException(); }
        public object LEN(object value) { throw new NotImplementedException(); }
        public object LENB(object value) { throw new NotImplementedException(); }
        public object LEFT(object value) { throw new NotImplementedException(); }
        public object LEFTB(object value) { throw new NotImplementedException(); }
        public object RIGHT(object value) { throw new NotImplementedException(); }
        public object RIGHTB(object value) { throw new NotImplementedException(); }
        public object REPLACE(object value) { throw new NotImplementedException(); }
        public object SPACE(object value) { throw new NotImplementedException(); }
        public object SPLIT(object value) { throw new NotImplementedException(); }
        public object STRCOMP(object string1, object string2) { return STRCOMP(string1, string2, 0); }
        public object STRCOMP(object string1, object string2, object compare) { return ToVBScriptNullable<int>(STRCOMP_Internal(string1, string2, compare)); }
        private int? STRCOMP_Internal(object string1, object string2, object compare)
        {
            throw new NotImplementedException();
        }
        public string TRIM(object value) { throw new NotImplementedException(); }
        public string LTRIM(object value) { throw new NotImplementedException(); }
        public string RTRIM(object value) { throw new NotImplementedException(); }
        public string LCASE(object value) { throw new NotImplementedException(); }
        public string UCASE(object value) { throw new NotImplementedException(); }
        // - Type comparisons
        public object ISARRAY(object value) { throw new NotImplementedException(); }
        public object ISDATE(object value) { throw new NotImplementedException(); }
        public object ISEMPTY(object value) { throw new NotImplementedException(); }
        public object ISNULL(object value) { throw new NotImplementedException(); }
        public object ISNUMERIC(object value) { throw new NotImplementedException(); }
        public object ISOBJECT(object value) { throw new NotImplementedException(); }
        public object TYPENAME(object value)
        {
            if (value == null)
                return "Null";
            if (value == DBNull.Value)
                return "Empty";
            if (base.IsVBScriptNothing(value))
                return "Nothing";
            var type = value.GetType();
            var sourceClassName = type.GetCustomAttributes(typeof(SourceClassName), inherit: true).FirstOrDefault() as SourceClassName;
            if (sourceClassName != null)
                return sourceClassName.Name;
            return value.GetType().Name; // TODO: Does this deal with COM objects such as "Recordset" or will it show "_COMObject" (or whatever)
        }
        public object VARTYPE(object value) { throw new NotImplementedException(); }
        // - Array functions
        public object ARRAY(object value) { throw new NotImplementedException(); }
        public object ERASE(object value) { throw new NotImplementedException(); }
        public object JOIN(object value) { throw new NotImplementedException(); }
        public object LBOUND(object value) { throw new NotImplementedException(); }
        public object UBOUND(object value) { throw new NotImplementedException(); }
        // - Date functions
        public DateTime NOW() { throw new NotImplementedException(); }
        public DateTime DATE() { throw new NotImplementedException(); }
        public DateTime TIME() { throw new NotImplementedException(); }
        public object DATEADD(object value) { throw new NotImplementedException(); }
        public object DATESERIAL(object year, object month, object date) { throw new NotImplementedException(); }
        public object DATEVALUE(object value) { throw new NotImplementedException(); }
        public object TIMESERIAL(object value) { throw new NotImplementedException(); }
        public object TIMEVALUE(object value) { throw new NotImplementedException(); }
        public object NOW(object value) { throw new NotImplementedException(); }
        public object DAY(object value) { throw new NotImplementedException(); }
        public object MONTH(object value) { throw new NotImplementedException(); }
        public object YEAR(object value) { throw new NotImplementedException(); }
        public object WEEKDAY(object value) { throw new NotImplementedException(); }
        public object HOUR(object value) { throw new NotImplementedException(); }
        public object MINUTE(object value) { throw new NotImplementedException(); }
        public object SECOND(object value) { throw new NotImplementedException(); }
        // - Object creation
        public object CREATEOBJECT(object value) { throw new NotImplementedException(); }
        public object GETOBJECT(object value) { throw new NotImplementedException(); }
        public object EVAL(object value) { throw new NotImplementedException(); }
        public object EXECUTE(object value) { throw new NotImplementedException(); }
        public object EXECUTEGLOBAL(object value) { throw new NotImplementedException(); }

        // Array definitions
        public void NEWARRAY(IEnumerable<object> dimensions, Action<object> targetSetter)
        {
            throw new NotImplementedException(); // TODO
        }

        public void RESIZEARRAY(object array, IEnumerable<object> dimensions, Action<object> targetSetter)
        {
            throw new NotImplementedException(); // TODO
        }

        private IEnumerable<int> GetDimensions(IEnumerable<object> dimensions)
        {
            if (dimensions == null)
                throw new ArgumentNullException("dimensions");

            throw new NotImplementedException(); // TODO
        }

        // TODO: Consider using error translations from http://blogs.msdn.com/b/ericlippert/archive/2004/08/25/error-handling-in-vbscript-part-three.aspx
        
        public void CLEARANYERROR() { } // TODO
        public void SETERROR(Exception e) { } // TODO

        public int GETERRORTRAPPINGTOKEN() { throw new NotImplementedException(); } // TODO
        public void RELEASEERRORTRAPPINGTOKEN(int token) { throw new NotImplementedException(); } // TODO

        public void STARTERRORTRAPPINGANDCLEARANYERROR(int token) { throw new NotImplementedException(); } // TODO
        public void STOPERRORTRAPPINGANDCLEARANYERROR(int token) { throw new NotImplementedException(); } // TODO

        public void HANDLEERROR(int token, Action action) { throw new NotImplementedException(); } // TODO

        public bool IF(Func<object> valueEvaluator, int errorToken)
        {
            if (valueEvaluator == null)
                throw new ArgumentNullException("valueEvaluator");

            throw new NotImplementedException(); // TODO
        }

        /// <summary>
        /// VBScript has comparisons that will return true, false or Null (meaning DBNull.Value) which is a return type that is difficult to represent
        /// without resorting to "object" (which could be anything) or an enum (which wouldn't be the end of the world). I think the best approach,
        /// though, is to return a nullable bool from methods internally and then translate this for VBScript (so null becomes DBNull.Value).
        /// The same approach works for other non-nullable types.
        /// </summary>
        private static object ToVBScriptNullable<T>(T? value) where T : struct
        {
            if (value == null)
                return DBNull.Value;
            return value.Value;
        }
    }
}
