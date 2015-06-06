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
        // These are going to be used repeatedly so let's calculate them once and reuse them
        private readonly static double MIN_DATE_VALUE_AS_DOUBLE = VBScriptConstants.EarliestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays;
        private readonly static double MAX_DATE_VALUE_AS_DOUBLE = VBScriptConstants.LatestPossibleDate.Subtract(VBScriptConstants.ZeroDate).TotalDays;

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
            // one-or-both-Null (Null), both-strings (concatenate) or one-string-with-Empty (return string). Note that single-Empty is not a simple case - for most values
            // it is (CInt(1) + Empty = CInt(1), for example) but when Empty is added to a Boolean then the type changes to an Integer.
            l = _valueRetriever.VAL(l);
            r = _valueRetriever.VAL(r);
            if ((l == DBNull.Value) || (r == DBNull.Value))
                return DBNull.Value;
            if ((l == null) && (r == null))
                return (Int16)0;
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
            if (((lCurrency != null) && (rDate != null)) || ((rCurrency != null) && (lDate != null)))
            {
                var currencyValue = lCurrency ?? rCurrency.Value;
                var dateValue = lDate ?? rDate.Value;
                var result = (double)currencyValue + dateValue.Subtract(VBScriptConstants.ZeroDate).TotalDays;
                if ((result >= MIN_DATE_VALUE_AS_DOUBLE) && (result <= MAX_DATE_VALUE_AS_DOUBLE))
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
                    // We know that one (and only one) of the values is a Currency and that the other is not a Date. We can retrieve the other value as a Double and
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
                if ((result >= MIN_DATE_VALUE_AS_DOUBLE) && (result <= MAX_DATE_VALUE_AS_DOUBLE))
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
                var overflowSafeResult = (Int16)((Int16)lByte.Value + (Int16)rByte.Value);
                if ((overflowSafeResult >= byte.MinValue) && (overflowSafeResult <= byte.MaxValue))
                    return (byte)overflowSafeResult;
                return overflowSafeResult;
            }
            else if ((lByte != null) && (rBoolean != null))
                return (Int16)(lByte.Value + (rBoolean.Value ? -1 : 0));
            else if ((rByte != null) && (lBoolean != null))
                return (Int16)(rByte.Value + (lBoolean.Value ? -1 : 0));
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
                    lInteger = lBoolean.Value ? (Int16)(-1) : (Int16)0;
                else if (lByte != null)
                    lInteger = (Int16)lByte.Value;
            }
            if (rInteger == null)
            {
                if (rBoolean != null)
                    rInteger = rBoolean.Value ? (Int16)(-1) : (Int16)0;
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
            else if (((lInteger != null) && (r == null)) || ((rInteger != null) && (l == null)))
                return lInteger ?? rInteger.Value;

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

        public object SUBT(object value)
        {
            value = _valueRetriever.VAL(value);
            if (value == null)
                return (Int16)0;
            else if (value == DBNull.Value)
                return DBNull.Value;

            // Booleans are not supported here ("-true" results in a "Type mismatch" error in VBScript)
            if (value is bool)
                throw new TypeMismatchException();

            // Force the value into a number (this will ensure that strings are parsed if they are numeric, but not if they're string representations of booleans
            // or dates, which aren't valid for this operation)
            value = _valueRetriever.NUM(value);

            // Bytes are easy - they're either zero, which is a no-op, or they need negating and returning as an Integer (Int16)
            var valueByte = TryToCoerceInto<byte>(value);
            if (valueByte != null)
            {
                if (valueByte.Value == 0)
                    return valueByte.Value;
                return (Int16)(-valueByte.Value);
            }

            // VBScript Integers (ie. Int16) are fairly simple - they are negated unless they are the one value that would overflow (-32768), in which case it will
            // become a Long (Int32)
            var valueInteger = TryToCoerceInto<Int16>(value);
            if (valueInteger != null)
            {
                if (valueInteger == Int16.MinValue)
                    return -valueInteger; // The minus operator will change the type to Int32 here (which is what we want)
                return (Int16)(-valueInteger);
            }

            // VBScript Longs (Int32) are basically the same as Integers
            var valueLong = TryToCoerceInto<Int32>(value);
            if (valueLong != null)
            {
                if (valueLong == Int32.MinValue)
                    return -((double)valueLong); // The minus operator will keep the type as Int32, which will overflow back to where we came from if we negate it - so cast to double first
                return (Int32)(-valueLong);
            }

            // Same sort of deal applies to Dates..
            var valueDate = TryToCoerceInto<DateTime>(value);
            if (valueDate != null)
            {
                var valueAsDouble = -valueDate.Value.Subtract(VBScriptConstants.ZeroDate).TotalDays;
                if ((valueAsDouble < MIN_DATE_VALUE_AS_DOUBLE) || (valueAsDouble > MAX_DATE_VALUE_AS_DOUBLE))
                    return valueAsDouble;
                return VBScriptConstants.ZeroDate.AddDays(valueAsDouble);
            }

            // Currency has no edge cases issues since the min Currency value = -(max Currency value)
            var valueCurrency = TryToCoerceInto<Decimal>(value);
            if (valueCurrency != null)
                return -valueCurrency.Value;

            // Fall back to a Double if all else fails
            return -AsDouble(value);
        }

        public object SUBT(object l, object r)
        {
            return ADD(l, SUBT(r));
        }

        public object MULT(object l, object r)
        {
            // Address simplest cases first - ensure both values are non-object references (or may be coerced into value types), then check for double-Empty (Integer zero),
            // one-or-both-Null (Null), both-strings (concatenate) or one-string-with-Empty (return string). Note that single-Empty is not a simple case - for most values
            // it is (CInt(1) + Empty = CInt(1), for example) but when Empty is added to a Boolean then the type changes to an Integer.
            l = _valueRetriever.VAL(l);
            r = _valueRetriever.VAL(r);
            if ((l == DBNull.Value) || (r == DBNull.Value))
                return DBNull.Value;
            if ((l == null) && (r == null))
                return (Int16)0;

            // Multiplying two Booleans always returns an Integer, so the smallest type is the Byte - a Byte is only returned if BOTH values are Bytes or if one is a Byte
            // and the other is Empty (in which case a zero of type Byte will be returned)
            // by Empty will be an Integer zero.
            var lByte = TryToCoerceInto<byte>(l);
            var rByte = TryToCoerceInto<byte>(r);
            if ((lByte != null) && (rByte != null))
            {
                var overflowSafeResult = (Int16)(lByte.Value * rByte.Value);
                if ((overflowSafeResult >= byte.MinValue) && (overflowSafeResult <= byte.MaxValue))
                    return (byte)overflowSafeResult;
                return overflowSafeResult;
            }
            if (((lByte != null) && (r == null)) || ((rByte != null) && (l == null)))
                return (byte)0;

            // Since multiplying Integers or Integers with Booleans or Integers with Bytes or Bytes with Booleans all result in an Integer being returned (unless an
            // overflow into a Long is required), we'll treat any Byte or Boolean values here as Integer. An Integer or Byte or Boolean multipled by Empty results
            // in an Integer zero.
            var lInteger = (lByte != null) ? ((Int16)lByte.Value) : TryToCoerceInto<Int16>(l);
            if (lInteger == null)
            {
                var lBoolean = TryToCoerceInto<bool>(l);
                if (lBoolean != null)
                    lInteger = lBoolean.Value ? (Int16)(-1) : (Int16)0;
            }
            var rInteger = (rByte != null) ? ((Int16)rByte.Value) : TryToCoerceInto<Int16>(r);
            if (rInteger == null)
            {
                var rBoolean = TryToCoerceInto<bool>(r);
                if (rBoolean != null)
                    rInteger = rBoolean.Value ? (Int16)(-1) : (Int16)0;
            }
            if ((lInteger != null) && (rInteger != null))
            {
                var overflowSafeResult = lInteger.Value * rInteger.Value;
                if ((overflowSafeResult >= Int16.MinValue) && (overflowSafeResult <= Int16.MaxValue))
                    return (Int16)overflowSafeResult;
                return overflowSafeResult;
            }
            if (((lInteger != null) && (r == null)) || ((rInteger != null) && (l == null)))
                return (Int16)0;

            // Long (aka Int32) is similar to Integer (Int16) - if both values are Longs or if one is Long and the other is a smaller type (Boolean, Byte, Integer) then
            // then Long, unless an overflow into Double is required
            var lLong = TryToCoerceInto<Int32>(l);
            var rLong = TryToCoerceInto<Int32>(r);
            if ((lLong == null) && (lInteger != null))
                lLong = (Int32)lInteger; // This will already have considered Boolean and Byte values where applicable (see above)
            if ((rLong == null) && (rInteger != null))
                rLong = (Int32)rInteger; // This will already have considered Boolean and Byte values where applicable (see above)
            if ((lLong != null) && (rLong != null))
            {
                var result = (Double)lLong.Value * (Double)rLong.Value;
                if ((result >= Int32.MinValue) && (result <= Int32.MaxValue))
                    return (Int32)result;
                return result;
            }
            if (((lLong != null) && (r == null)) || ((rLong != null) && (l == null)))
                return 0;

            // Two Currency value multiplied will result in a Currency. If the result is an overflow then an error is raised, rather than moving up to a Double. A Currency
            // multiplied by a Date or Double results in a Double, but any other type results in a Currency.
            var lCurrency = TryToCoerceInto<Decimal>(l);
            var rCurrency = TryToCoerceInto<Decimal>(r);
            if ((lCurrency != null) || (rCurrency != null))
            {
                Tuple<Decimal, Decimal> currencyValuesToMultiplyIfAvailable;
                if ((lCurrency != null) && (rCurrency != null))
                    currencyValuesToMultiplyIfAvailable = Tuple.Create(lCurrency.Value, rCurrency.Value);
                else if ((lCurrency != null) && (rLong != null))
                    currencyValuesToMultiplyIfAvailable = Tuple.Create(lCurrency.Value, (Decimal)rLong.Value);
                else if ((rCurrency != null) && (lLong != null))
                    currencyValuesToMultiplyIfAvailable = Tuple.Create(rCurrency.Value, (Decimal)lLong.Value);
                else if ((l == null) || (r == null))
                {
                    // We know that one of the value is a Currency so if either of the values is null then we're combining a Currency with null (Empty in VBScript)
                    return 0m;
                }
                else
                    currencyValuesToMultiplyIfAvailable = null;
                if (currencyValuesToMultiplyIfAvailable != null)
                {
                    Decimal result;
                    try
                    {
                        result = currencyValuesToMultiplyIfAvailable.Item1 * currencyValuesToMultiplyIfAvailable.Item2;
                    }
                    catch (OverflowException e)
                    {
                        throw new VBScriptOverflowException(e);
                    }
                    if ((result < VBScriptConstants.MinCurrencyValue) || (result > VBScriptConstants.MaxCurrencyValue))
                        throw new VBScriptOverflowException();
                    return result;
                }
            }

            // Multiplying Dates or multiplying anything BY a Date will result in a Double. Might as well wrap this into a final treat-as-Double if nothing else matched
            // fallback (this covers numeric string cases too)
            var lDate = TryToCoerceInto<DateTime>(l);
            var rDate = TryToCoerceInto<DateTime>(r);
            var lDouble = (lDate != null) ? lDate.Value.Subtract(VBScriptConstants.ZeroDate).TotalDays : AsDouble(l);
            var rDouble = (rDate != null) ? rDate.Value.Subtract(VBScriptConstants.ZeroDate).TotalDays : AsDouble(r);
            return lDouble * rDouble;
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

        public object MOD(object l, object r)
        {
            // Note: Null values trump division-by zero (so check them first) but overflow trumps division-by-zero, so we can't throw as soon as we find that "r" is zero,
            // we need to continue until we confirm that "l" does not overflow. With the exception of double-Empty; there's no chance of an overflow then, it's definitely
            // safe to division-by-zero it!
            l = _valueRetriever.VAL(l);
            r = _valueRetriever.VAL(r);
            if ((l == DBNull.Value) || (r == DBNull.Value))
                return DBNull.Value;
            if ((l == null) && (r == null))
                throw new VBScriptDivisionByZeroException();

            var rByte = TryToCoerceInto<byte>(r);
            if ((l == null) && (rByte != null))
                return (Int16)0; // Empty gets treated as an Integer, so the combined result is an Integer
            var lByte = TryToCoerceInto<byte>(l);
            if ((lByte != null) && (rByte != null))
            {
                if (rByte == 0)
                    throw new VBScriptDivisionByZeroException();
                return (byte)(lByte % rByte);
            }

            var rInteger = (rByte != null) ? (Int16?)rByte.Value : TryToCoerceInto<Int16>(r);
            if ((rInteger == null) && (r != null))
            {
                var rBoolean = TryToCoerceInto<bool>(r);
                if (rBoolean != null)
                    rInteger = rBoolean.Value ? (Int16)(-1) : (Int16)0;
            }
            if ((l == null) && (rInteger != null))
                return (Int16)0;
            var lInteger = (lByte != null) ? (Int16?)lByte.Value : TryToCoerceInto<Int16>(l);
            if ((lInteger == null) && (l != null))
            {
                var lBoolean = TryToCoerceInto<bool>(l);
                if (lBoolean != null)
                    lInteger = lBoolean.Value ? (Int16)(-1) : (Int16)0;
            }
            if ((lInteger != null) && (rInteger != null))
            {
                if (rInteger == 0)
                    throw new VBScriptDivisionByZeroException();
                return (Int16)(lInteger.Value % rInteger.Value);
            }

            var rLong = (rInteger != null) ? (int?)rInteger.Value : TryToCoerceInto<Int32>(r);
            if ((rLong == null) && (r != null))
            {
                try
                {
                    var rDate = TryToCoerceInto<DateTime>(r);
                    if (rDate != null)
                        rLong = Convert.ToInt32(rDate.Value.Subtract(VBScriptConstants.ZeroDate).TotalDays);
                    else
                        rLong = Convert.ToInt32(AsDouble(r));
                }
                catch (OverflowException e)
                {
                    throw new VBScriptOverflowException(e);
                }
            }
            if ((l == null) && (rLong != null))
                return (Int32)0;
            var lLong = (lInteger != null) ? (int?)lInteger.Value : TryToCoerceInto<Int32>(l);
            if ((lLong == null) && (l != null))
            {
                try
                {
                    var lDate = TryToCoerceInto<DateTime>(l);
                    if (lDate != null)
                        lLong = Convert.ToInt32(lDate.Value.Subtract(VBScriptConstants.ZeroDate).TotalDays);
                    else
                        lLong = Convert.ToInt32(AsDouble(l));
                }
                catch (OverflowException e)
                {
                    throw new VBScriptOverflowException(e);
                }
            }
            if ((r == null) || (rLong == 0))
                throw new VBScriptDivisionByZeroException();
            return lLong % rLong;
        }

        private double AsDouble(object value)
        {
            // The rules that the NUM function must abide by are the same as the ones we want applied here - try to interpret a value as a string by ensuring it is
            // a value type (requiring a default parameterless member if not a value type, otherwise an exception will be raised) and then checking for the already-
            // numeric-esque types (Boolean, Integer, Date, etc..) and allowing some flexibility (strings are allowed if they are numeric, but not if they are string
            // representations of boolean or date values). Null is not acceptable but Empty is.
            if (value is double)
                return (double)value;
            var numericValue = _valueRetriever.NUM(value);
            if (numericValue is double)
                return (double)numericValue;
            return Convert.ToDouble(numericValue);
        }

        private T? TryToCoerceInto<T>(object value) where T : struct
        {
            if (value is T)
                return (T)value;
            return (T?)null;
        }
    }
}