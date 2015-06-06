using System;
using CSharpSupport.Exceptions;

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
            // While there is information about this operation at https://msdn.microsoft.com/en-us/library/kd1e4aey(v=vs.84).aspx, it is lacking on details such as what
            // happens when both values are numbers (or dates) of different types (or what happens if they're the same type but their addition would result in an overflow).
            // See the test suite for further details.
            
            // Address simplest cases first - ensure both values are non-object references (or may be coerced into value types), then check for double-Empty (Integer zero),
            // one-or-both-Null (Null), both-strings (concatenate) or one-string-with-Empty (return string).
            l = _valueRetriever.VAL(l);
            r = _valueRetriever.VAL(r);
            if ((l == DBNull.Value) || (r == DBNull.Value))
                return DBNull.Value;
            if ((l == null) && (r == null))
                return (Int16)0;
            else if ((l == null) || (r == null))
                return l ?? r;
            var lString = l as string;
            var rString = r as string;
            if ((lString != null) && (rString != null))
                return lString + rString;
            else if (((lString != null) && (r == null)) || ((rString != null) && (l == null)))
                return lString ?? rString;

            // The most unusual case is worth addressing next - Currency almost never changes its type, it will overflow if the result of the operation is outside of the
            // range that Currency can describe (whereas other types will move up to the next biggest type - Currency COULD do this with Double, but doesn't). The notable
            // exception is that if a Currency is added to a Date then the result will be a date.. unless the result would overflow the expressible Date range, in which
            // case it will be a Double.
            var lCurrency = TryToCoerceInto<decimal>(l);
            var rCurrency = TryToCoerceInto<decimal>(r);
            var lDate = TryToCoerceInto<DateTime>(l);
            var rDate = TryToCoerceInto<DateTime>(r);
            var minDateValueAsDouble = VBScriptConstants.EarliestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays;
            var maxDateValueAsDouble = VBScriptConstants.LatestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays;
            if (((lCurrency != null) && (rDate != null)) || ((rCurrency != null) && (lDate != null)))
            {
                var currencyValue = lCurrency ?? rCurrency.Value;
                var dateValue = lDate ?? rDate.Value;
                var result = (double)currencyValue + dateValue.Subtract(VBScriptConstants.ZeroDate).TotalDays;
                if ((result >= minDateValueAsDouble) && (result <= maxDateValueAsDouble))
                    return VBScriptConstants.ZeroDate.AddDays(result);
                return result;
            }
            if ((lCurrency != null) || (rCurrency != null))
            {
                Decimal firstCurrencyValue, secondCurrencyValue;
                if ((lCurrency != null) && (rCurrency != null))
                {
                    firstCurrencyValue = lCurrency.Value;
                    secondCurrencyValue = rCurrency.Value;
                }
                else
                {
                    // We know that one (and only one) of the values if a Currency and that the other is not a Date. We can retrieve the other value as a Double and
                    // then perform the translation (there is no way that precision can be lost since only Currency has different precision rules from Double and we
                    // know that the "other value" is not a Currency).
                    firstCurrencyValue = lCurrency ?? rCurrency.Value;
                    try
                    {
                        secondCurrencyValue = Convert.ToDecimal(AsDouble(lCurrency == null ? l : r));
                    }
                    catch (Exception e)
                    {
                        throw new VBScriptOverflowException(e);
                    }
                }
                Decimal result;
                try
                {
                    result = firstCurrencyValue + secondCurrencyValue;
                }
                catch (OverflowException e)
                {
                    throw new VBScriptOverflowException(e);
                }
                if ((result < VBScriptConstants.MinCurrencyValue) || (result > VBScriptConstants.MaxCurrencyValue))
                    throw new VBScriptOverflowException();
                return result;
            }

            // Dates are the other unusual case - operations involving a Date also return a Date in most cases, though they will overflow into a Double if required
            // (unlike Currency, which will throw an overflow exception rather than move up to a Double)
            if ((lDate != null) || (rDate != null))
            {
                var firstDoubleValue = (lDate != null) ? lDate.Value.Subtract(VBScriptConstants.ZeroDate).TotalDays : AsDouble(l);
                var secondDoubleValue = (rDate != null) ? rDate.Value.Subtract(VBScriptConstants.ZeroDate).TotalDays : AsDouble(r);
                var result = firstDoubleValue + secondDoubleValue;
                if ((result >= minDateValueAsDouble) && (result <= maxDateValueAsDouble))
                    return VBScriptConstants.ZeroDate.AddDays(result);
                return result;
            }

            // The simplest case is both-Booleans
            var lBoolean = TryToCoerceInto<bool>(l);
            var rBoolean = TryToCoerceInto<bool>(r);
            if ((lBoolean != null) && (rBoolean != null))
            {
                if (!lBoolean.Value && !rBoolean.Value)
                    return false;
                if (lBoolean.Value && rBoolean.Value)
                    return (Int16)(-2); // The two true values are treated as -1 for arithmetic and the -2 result overflows Boolean into a VBScript Integer
                return true; // If one is false and one true then it's -1 + 0 = -1 = true
            }

            // The next smallest type is Byte, which would make it simple except that the case of Byte + Boolean is special since that must overflow into Integer (since the
            // type Byte, which is 0-255, can not contain all values of Boolean, which are 0 or -1)
            var lByte = TryToCoerceInto<byte>(l);
            var rByte = TryToCoerceInto<byte>(r);
            if ((lByte != null) && (rByte != null))
            {
                var overflowSafeResult = (Int16)lByte.Value + (Int16)rByte.Value;
                if ((overflowSafeResult >= byte.MinValue) && (overflowSafeResult <= byte.MaxValue))
                    return (byte)overflowSafeResult;
                return overflowSafeResult;
            }
            else if ((lByte != null) && (rBoolean != null))
                return (Int16)lByte.Value + (rBoolean.Value ? (Int16)(-1) : (Int16)0);
            else if ((rByte != null) && (lBoolean != null))
                return (Int16)rByte.Value + (lBoolean.Value ? (Int16)(-1) : (Int16)0);
            else if ((lByte != null) && (r == null))
                return lByte.Value;
            else if ((rByte != null) && (l == null))
                return rByte.Value;

            // Next up is VBScript's Integer - aka C#'s Int16. If both values are Integer then the result will be Integer, unless it would overflow, in which case
            // it will become a Long (Int32). If one value is an Integer and the other is Empty then it's a no-op. If One value is an Integer and the other a
            // Boolean or Byte, then they will be promoted to an Integer and then it's a two-Integer operation.
            var lInteger = TryToCoerceInto<Int16>(l);
            var rInteger = TryToCoerceInto<Int16>(r);
            if (((lInteger != null) && (r == null)) || ((rInteger != null) && (l == null)))
                return lInteger ?? rInteger.Value;
            if (lInteger == null)
            {
                if (lBoolean != null)
                    lInteger = lBoolean.Value ? (Int16)0 : (Int16)(-1);
                else if (lByte != null)
                    lInteger = (Int16)lByte.Value;
            }
            if (rInteger == null)
            {
                if (rBoolean != null)
                    rInteger = rBoolean.Value ? (Int16)0 : (Int16)(-1);
                else if (rByte != null)
                    rInteger = (Int16)rByte.Value;
            }
            if ((lInteger != null) && (rInteger != null))
            {
                var result = (Int32)lInteger.Value + (Int32)rInteger.Value;
                if ((result >= Int16.MinValue) && (result <= Int16.MaxValue))
                    return (Int16)result;
                return result;
            }

            // Long (aka Int32) is handled in the same manner similar as Integer (Int16), it will overflow into Double if required
            var lLong = TryToCoerceInto<Int32>(l);
            var rLong = TryToCoerceInto<Int32>(r);
            if (((lLong != null) && (r == null)) || ((rLong != null) && (l == null)))
                return lLong ?? rLong.Value;
            if ((lLong == null) && (lInteger != null))
                lLong = (Int32)lInteger; // This will already have considered lBoolean and lByte values where applicable (see above)
            if ((rLong == null) && (rInteger != null))
                rLong = (Int32)rInteger; // This will already have considered rBoolean and rByte values where applicable (see above)
            if ((lLong != null) && (rLong != null))
            {
                var result = (Double)lLong.Value + (Double)rLong.Value;
                if ((result >= Int32.MinValue) && (result <= Int32.MaxValue))
                    return (Int32)result;
                return result;
            }

            // Now the only option open to us is to treat both values as Double. This covers cases where one or both of the values is a string (which will be
            // parsed into Doubles if numeric - note that boolean and date string representations are not acceptable).
            return AsDouble(l) + AsDouble(r);
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



        private double AsDouble(object value)
        {
            // The rules that the NUM function must abide by are the same as the ones we want applied here - try to interpret a value as a string by ensuring it is
            // a value type (requiring a default parameterless member if not a value type, otherwise an exception will be raised) and then checking for the already-
            // numeric-esque types (Boolean, Integer, Date, etc..) and allowing some flexibility (strings are allowed if they are numeric, but not if they are string
            // representations of boolean or date values). Null is not acceptable but Empty is.
            return Convert.ToDouble(_valueRetriever.NUM(value));
        }

        private T? TryToCoerceInto<T>(object value) where T : struct
        {
            if (value is T)
                return (T)value;
            return (T?)null;
        }
    }
}
