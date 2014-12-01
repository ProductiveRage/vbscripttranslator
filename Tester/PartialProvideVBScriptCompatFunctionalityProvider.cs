using CSharpSupport;
using CSharpSupport.Implementations;
using System;
using System.Collections.Generic;

namespace Tester
{
    /// <summary>
    /// This is intended to be built up over time and used with classes that are output by the translator. Clearly, at this point, it is noticeably lacking
    /// in working functionality, but it provides something, at least, to use to test the very simple programs that can be translated at this time.
    /// </summary>
    public class PartialProvideVBScriptCompatFunctionalityProvider : VBScriptEsqueValueRetriever, IProvideVBScriptCompatFunctionality
    {
        public PartialProvideVBScriptCompatFunctionalityProvider(Func<string, string> nameRewriter) : base(nameRewriter)
        {
            Constants = new VBScriptConstants();
        }

        public VBScriptConstants Constants { get; private set; }

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

        // Logical operators
        public int NOT(object o) { throw new NotImplementedException(); }
        public int AND(object l, object r) { throw new NotImplementedException(); }
        public int OR(object l, object r) { throw new NotImplementedException(); }
        public int XOR(object l, object r) { throw new NotImplementedException(); }

        // Comparison operators
        public int EQ(object l, object r) { throw new NotImplementedException(); }
        public int NOTEQ(object l, object r) { throw new NotImplementedException(); }
        public int LT(object l, object r) { throw new NotImplementedException(); }
        public int GT(object l, object r) { throw new NotImplementedException(); }
        public int LTE(object l, object r) { throw new NotImplementedException(); }
        public int GTE(object l, object r) { throw new NotImplementedException(); }
        public int IS(object l, object r) { throw new NotImplementedException(); }
        public int EQV(object l, object r) { throw new NotImplementedException(); }
        public int IMP(object l, object r) { throw new NotImplementedException(); }

        // Builtin functions - TODO: These are not fully specified yet (eg. LEFT requires more than one parameter and INSTR requires multiple parameters and
        // overloads to deal with optional parameters)
        // - Type conversions
        public object CBYTE(object value) { throw new NotImplementedException(); }
        public object CBOOL(object value) { throw new NotImplementedException(); }
        public object CDBL(object value) { throw new NotImplementedException(); }
        public object CDATE(object value) { throw new NotImplementedException(); }
        public object CINT(object value) { throw new NotImplementedException(); }
        public object CLNG(object value) { throw new NotImplementedException(); }
        public object CSNG(object value) { throw new NotImplementedException(); }
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
        public object STRCOMP(object value) { throw new NotImplementedException(); }
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
        public object TYPENAME(object value) { throw new NotImplementedException(); }
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
        public object DATESERIAL(object value) { throw new NotImplementedException(); }
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

        public int GETERRORTRAPPINGTOKEN() { throw new NotImplementedException(); } // TODO
        public void RELEASEERRORTRAPPINGTOKEN(int token) { throw new NotImplementedException(); } // TODO

        public void STARTERRORTRAPPING(int token) { throw new NotImplementedException(); } // TODO
        public void STOPERRORTRAPPING(int token) { throw new NotImplementedException(); } // TODO

        public void HANDLEERROR(int token, Action action) { throw new NotImplementedException(); } // TODO

        public bool IF(Func<object> valueEvaluator, int errorToken)
        {
            if (valueEvaluator == null)
                throw new ArgumentNullException("valueEvaluator");

            throw new NotImplementedException(); // TODO
        }
    }
}
