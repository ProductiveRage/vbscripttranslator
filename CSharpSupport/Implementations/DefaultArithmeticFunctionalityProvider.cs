using System;

namespace CSharpSupport.Implementations
{
    /// <summary>
    /// This handles a subset of IProvideVBScriptCompatFunctionalityToIndividualRequests - the arithmetic functions. In VBScript, these are no simple task due to the wide
    /// variety of behaviours that are exhibited, depending upon the type of the value(s) being manipulated.
    /// </summary>
    public class DefaultArithmeticFunctionalityProvider
    {
        private readonly IAccessValuesUsingVBScriptRules _valueRetriever;
        public DefaultArithmeticFunctionalityProvider(IAccessValuesUsingVBScriptRules valueRetriever)
        {
            if (valueRetriever == null)
                throw new ArgumentNullException("valueRetriever");

            _valueRetriever = valueRetriever;
        }

        public object ADD(object l, object r)
        {
            // See https://msdn.microsoft.com/en-us/library/kd1e4aey(v=vs.84).aspx
            l = _valueRetriever.VAL(l);
            r = _valueRetriever.VAL(r);
            if ((l == DBNull.Value) || (r == DBNull.Value))
                return DBNull.Value;
            if ((l == null) && (r == null))
                return (Int16)0;
            else if ((l == null) || (r == null))
                return l ?? r;
            if ((l is string) && (r is string))
                return (string)l + (string)r;

            // TODO: This just covers some simple cases, it is not the full implementation yet
            l = _valueRetriever.NUM(l, r);
            r = _valueRetriever.NUM(r, l);
            if ((l is short) && (r is short))
            {
                var shortL = (short)l;
                var shortR = (short)r;
                if (shortR >= 0)
                {
                    if (shortL <= short.MaxValue - shortR)
                        return (short)(shortL + shortR);
                    return shortL + shortR; // Note: short + short will return int so we don't need any casts here
                }
            }

            /*
            // TODO: Need to move into new data type if would overflow the current one

            // Now we need to treat both as numbers. We need them both to be a consistent type so that we can perform the addition. Using the
            // numericValuesTheTypeMustBeAbleToContain method signature of NUM will allow us  to do that, but we'll still have to inspect the
            // types to add correctly (NUM will only return VBScript-supported numeric types, though, which will make this easier).
            l = NUM(l, r);
            r = NUM(r, l);
            if (l is DateTime)
            {
                var dateResult = ((DateTime)l).AddDays(DateToDouble((DateTime)r));
                if ((dateResult < VBScriptConstants.EarliestPossibleDate) || (dateResult > VBScriptConstants.LatestPossibleDate))
                    throw new VBScriptOverflowException();
                return dateResult;
            }
            else if (l is decimal)
            {
                var decimalL = (decimal)l;
                var decimalR = (decimal)r;
                if ((decimalR > 0) && (VBScriptConstants.MaxCurrencyValue - decimalR > decimalL)) // TODO: Ensure include test coverage for this!
                    throw new VBScriptOverflowException();
                else if ((decimalR < 0) && (decimalL < VBScriptConstants.MinCurrencyValue - decimalR)) // TODO: Ensure include test coverage for this!
                    throw new VBScriptOverflowException();
                return decimalL + decimalR;
            }
            //return NUM(l) + NUM(r);
             */
            throw new NotImplementedException(); // TODO
        }

        public object SUBT(object o)
        {
            throw new NotImplementedException(); // TODO
        }

        public object SUBT(object l, object r)
        {
            throw new NotImplementedException(); // TODO
        }

        public double MULT(object l, object r)
        {
            throw new NotImplementedException(); // TODO
        }

        public double DIV(object l, object r)
        {
            throw new NotImplementedException(); // TODO
        }

        public int INTDIV(object l, object r)
        {
            throw new NotImplementedException(); // TODO
        }

        public double POW(object l, object r)
        {
            throw new NotImplementedException(); // TODO
        }

        public double MOD(object l, object r)
        {
            throw new NotImplementedException(); // TODO
        }
    }
}
