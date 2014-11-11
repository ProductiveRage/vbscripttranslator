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

        public void GETERRORTRAPPINGTOKEN() { throw new NotImplementedException(); } // TODO
        public void RELEASEERRORTRAPPINGTOKEN(int token) { throw new NotImplementedException(); } // TODO

        public void STARTERRORTRAPPING(int token) { throw new NotImplementedException(); } // TODO
        public void STOPERRORTRAPPING(int token) { throw new NotImplementedException(); } // TODO

        public void HANDLEERROR(Action action, int token) { throw new NotImplementedException(); } // TODO

        public bool IF(Func<object> valueEvaluator, int errorToken)
        {
            if (valueEvaluator == null)
                throw new ArgumentNullException("valueEvaluator");

            throw new NotImplementedException(); // TODO
        }
    }
}
