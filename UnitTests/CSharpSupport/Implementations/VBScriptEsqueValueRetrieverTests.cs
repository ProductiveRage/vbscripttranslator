using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Implementations;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    public class VBScriptEsqueValueRetrieverTests
    {
        [Fact]
        public void VALDoesNotAlterMinusOneNull()
        {
            Assert.Equal(
                null,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(null)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOne()
        {
            Assert.Equal(
                1,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(1)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOnePointOneFloat()
        {
            Assert.Equal(
                1.1f,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(1.1f)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOnePointOneDouble()
        {
            Assert.Equal(
                1.1d,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(1.1d)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOnePointOneDecimal()
        {
            Assert.Equal(
                1.1m,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(1.1m)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOne()
        {
            Assert.Equal(
                -1,
                (new VBScriptEsqueValueRetriever(name => name)).VAL(-1)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusEmptyString()
        {
            Assert.Equal(
                "",
                (new VBScriptEsqueValueRetriever(name => name)).VAL("")
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusNonEmptyString()
        {
            Assert.Equal(
                "Test",
                (new VBScriptEsqueValueRetriever(name => name)).VAL("Test")
            );
        }

        [Fact]
        public void IFOfNullIsFalse()
        {
            Assert.False(
                (new VBScriptEsqueValueRetriever(name => name)).IF(null)
            );
        }

        [Fact]
        public void IFOfZeroIsFalse()
        {
            Assert.False(
                (new VBScriptEsqueValueRetriever(name => name)).IF(0)
            );
        }

        [Fact]
        public void IFOfOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF(1)
            );
        }

        [Fact]
        public void IFOfMinusOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF(-1)
            );
        }

        [Fact]
        public void IFOfOnePointOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF(1.1)
            );
        }

        /// <summary>
        /// VBScript doesn't round the number down to zero and find 0.1 to be false, it just checks that the number is non-zero
        /// </summary>
        [Fact]
        public void IFOfPointOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF(0.1)
            );
        }

        [Fact]
        public void IFOfPointStringRepresentationOfOneIsTrue()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF("1")
            );
        }

        [Fact]
        public void IFThrowsExceptionForBlanksString()
        {
            Assert.Throws<ArgumentException>(() =>
            {
                (new VBScriptEsqueValueRetriever(name => name)).IF("");
            });
        }

        [Fact]
        public void IFThrowsExceptionForNonNumericString()
        {
            Assert.Throws<ArgumentException>(() =>
            {
                (new VBScriptEsqueValueRetriever(name => name)).IF("one");
            });
        }

        [Fact]
        public void IFIgnoresWhiteSpaceWhenParsingStrings()
        {
            Assert.True(
                (new VBScriptEsqueValueRetriever(name => name)).IF("   1    ")
            );
        }

        [Fact]
        public void NativeClassSupportsMethodCallWithArguments()
        {
            Assert.Equal(
                new PseudoField { value = "value:F1" },
                (new VBScriptEsqueValueRetriever(name => name)).CALL(
                    new PseudoRecordset(),
                    new[] { "fields" },
                    new MultipleByValArgumentProvider("F1")
                ),
                new PseudoFieldObjectComparer()
            );
        }

        [Fact]
        public void NativeClassSupportsDefaultMethodCallWithArguments()
        {
            Assert.Equal(
                new PseudoField { value = "value:F1" },
                (new VBScriptEsqueValueRetriever(name => name)).CALL(
                    new PseudoRecordset(),
                    new string[0],
                    new MultipleByValArgumentProvider("F1")
                ),
                new PseudoFieldObjectComparer()
            );
        }

        [Fact]
        public void ADORecordsetSupportsNamedFieldAccess()
        {
            var recordset = new ADODB.Recordset();
            recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
            recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
            recordset.AddNew();
            recordset.Fields["name"].Value = "TestName";
            recordset.Update();

            Assert.Equal(
                recordset.Fields["name"],
                (new VBScriptEsqueValueRetriever(name => name)).CALL(
                    recordset,
                    new[] { "fields" },
                    new MultipleByValArgumentProvider("name")
                ),
                new ADOFieldObjectComparer()
            );
        }

        [Fact]
        public void ADORecordsetSupportsDefaultFieldAccess()
        {
            var recordset = new ADODB.Recordset();
            recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
            recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
            recordset.AddNew();
            recordset.Fields["name"].Value = "TestName";
            recordset.Update();

            Assert.Equal(
                recordset.Fields["name"],
                (new VBScriptEsqueValueRetriever(name => name)).CALL(
                    recordset,
                    new string[0],
                    new MultipleByValArgumentProvider("name")
                ),
                new ADOFieldObjectComparer()
            );
        }

        /// <summary>
        /// This describes an extremely common VBScript pattern - rstResults("name") needs to return a value by using the default Fields access
        /// (passing through "name" to it) and then default access of the Field's Value property (which requires a VAL call, which must be
        /// included in translated code if a value type is expected - any time other than when a SET statement is present as part of a
        /// variable assignment)
        /// </summary>
        [Fact]
        public void ADORecordsetSupportsDefaultFieldValueAccess()
        {
            var recordset = new ADODB.Recordset();
            recordset.Fields.Append("name", ADODB.DataTypeEnum.adVarChar, 20, ADODB.FieldAttributeEnum.adFldUpdatable);
            recordset.Open(CursorType: ADODB.CursorTypeEnum.adOpenUnspecified, LockType: ADODB.LockTypeEnum.adLockUnspecified, Options: 0);
            recordset.AddNew();
            recordset.Fields["name"].Value = "TestName";
            recordset.Update();

            var valueRetriever = new VBScriptEsqueValueRetriever(name => name);
            Assert.Equal(
                "TestName",
                valueRetriever.VAL(
                    valueRetriever.CALL(
                        recordset,
                        new string[0],
                        new MultipleByValArgumentProvider("name")
                    )
                )
            );
        }

        [Fact]
        public void OneDimensionalArrayAccessIsSupported()
        {
            var data = new object[] { "One" };
            Assert.Equal(
                "One",
                (new VBScriptEsqueValueRetriever(name => name)).CALL(
                    data,
                    new string[0],
                    new MultipleByValArgumentProvider()
                )
            );
        }

        private class ADOFieldObjectComparer : IEqualityComparer<object>
        {
            public new bool Equals(object x, object y)
            {
                if (x == null)
                    throw new ArgumentNullException("x");
                if (y == null)
                    throw new ArgumentNullException("y");
                var fieldX = x as ADODB.Field;
                if (fieldX == null)
                    throw new ArgumentException("x is not an ADODB.Field");
                var fieldY = y as ADODB.Field;
                if (fieldY == null)
                    throw new ArgumentException("y is not an ADODB.Field");
                return fieldX.Value == fieldY.Value;
            }

            public int GetHashCode(object obj)
            {
                return 0;
            }
        }

        private class PseudoFieldObjectComparer : IEqualityComparer<object>
        {
            public new bool Equals(object x, object y)
            {
                if (x == null)
                    throw new ArgumentNullException("x");
                if (y == null)
                    throw new ArgumentNullException("y");
                var fieldX = x as PseudoField;
                if (fieldX == null)
                    throw new ArgumentException("x is not a PseudoField");
                var fieldY = y as PseudoField;
                if (fieldY == null)
                    throw new ArgumentException("y is not a PseudoField");
                return (fieldX.value as string) == (fieldY.value as string);
            }

            public int GetHashCode(object obj)
            {
                return 0;
            }
        }

        private class PseudoRecordset
        {
            [IsDefault]
            public object fields(string fieldName)
            {
                return new PseudoField { value = "value:" + fieldName };
            }
        }

        private class PseudoField
        {
            [IsDefault]
            public object value { get; set; }
        }

        private class MultipleByValArgumentProvider : IProvideCallArguments
        {
            private readonly object[] _values;
            public MultipleByValArgumentProvider(params object[] values)
            {
                if (values == null)
                    throw new ArgumentNullException("values");

                _values = values.ToArray();
            }
            
            public int NumberOfArguments { get { return _values.Length; } }

            public IEnumerable<object> GetInitialValues()
            {
                return _values.ToArray();
            }

            public void OverwriteValueIfByRef(int index, object value)
            {
                // Since these  values are all ByVal we only need to perform validation, not do any actual value updating
                if ((index < 0) || (index >= _values.Length))
                    throw new ArgumentOutOfRangeException("index");
            }
        }
    }
}
