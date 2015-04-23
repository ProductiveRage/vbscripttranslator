using CSharpSupport;
using CSharpSupport.Attributes;
using CSharpSupport.Exceptions;
using CSharpSupport.Implementations;
using System;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace VBScriptTranslator.UnitTests.CSharpSupport.Implementations
{
    // TODO: Add a test that complements ExpressionGeneratorTests.PropertyAccessOnStringLiteralResultsInRuntimeError, that confirms that if a string value is
    // provided as call target that an error will be raised for property or method access attempts
    public class VBScriptEsqueValueRetrieverTests
    {
        [Fact]
        public void VALDoesNotAlterMinusOneNull()
        {
            Assert.Equal(
                null,
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(null)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOne()
        {
            Assert.Equal(
                1,
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(1)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOnePointOneFloat()
        {
            Assert.Equal(
                1.1f,
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(1.1f)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOnePointOneDouble()
        {
            Assert.Equal(
                1.1d,
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(1.1d)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOneOnePointOneDecimal()
        {
            Assert.Equal(
                1.1m,
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(1.1m)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusOne()
        {
            Assert.Equal(
                -1,
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL(-1)
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusEmptyString()
        {
            Assert.Equal(
                "",
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL("")
            );
        }

        [Fact]
        public void VALDoesNotAlterMinusNonEmptyString()
        {
            Assert.Equal(
                "Test",
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.VAL("Test")
            );
        }

        [Fact]
        public void IFOfNullIsFalse()
        {
            Assert.False(
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(null)
            );
        }

        [Fact]
        public void IFOfZeroIsFalse()
        {
            Assert.False(
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(0)
            );
        }

        [Fact]
        public void IFOfOneIsTrue()
        {
            Assert.True(
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(1)
            );
        }

        [Fact]
        public void IFOfMinusOneIsTrue()
        {
            Assert.True(
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(-1)
            );
        }

        [Fact]
        public void IFOfOnePointOneIsTrue()
        {
            Assert.True(
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(1.1)
            );
        }

        /// <summary>
        /// VBScript doesn't round the number down to zero and find 0.1 to be false, it just checks that the number is non-zero
        /// </summary>
        [Fact]
        public void IFOfPointOneIsTrue()
        {
            Assert.True(
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF(0.1)
            );
        }

        [Fact]
        public void IFOfPointStringRepresentationOfOneIsTrue()
        {
            Assert.True(
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF("1")
            );
        }

        [Fact]
        public void IFThrowsExceptionForBlanksString()
        {
            Assert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF("");
            });
        }

        [Fact]
        public void IFThrowsExceptionForNonNumericString()
        {
            Assert.Throws<TypeMismatchException>(() =>
            {
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF("one");
            });
        }

        [Fact]
        public void IFIgnoresWhiteSpaceWhenParsingStrings()
        {
            Assert.True(
                DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever.IF("   1    ")
            );
        }

        [Fact]
        public void NativeClassSupportsMethodCallWithArguments()
        {
            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            Assert.Equal(
                new PseudoField { value = "value:F1" },
                _.CALL(
                    new PseudoRecordset(),
                    new[] { "fields" },
                    _.ARGS.Val("F1")
                ),
                new PseudoFieldObjectComparer()
            );
        }

        [Fact]
        public void NativeClassSupportsDefaultMethodCallWithArguments()
        {
            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            Assert.Equal(
                new PseudoField { value = "value:F1" },
                _.CALL(
                    new PseudoRecordset(),
                    new string[0],
                    _.ARGS.Val("F1")
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

            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            Assert.Equal(
                recordset.Fields["name"],
                _.CALL(
                    recordset,
                    new[] { "fields" },
                    _.ARGS.Val("name")
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

            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            Assert.Equal(
                recordset.Fields["name"],
                _.CALL(
                    recordset,
                    new string[0],
                    _.ARGS.Val("name")
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

            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            Assert.Equal(
                "TestName",
                _.VAL(
                    _.CALL(
                        recordset,
                        new string[0],
                        _.ARGS.Val("name")
                    )
                )
            );
        }

        [Fact]
        public void OneDimensionalArrayAccessIsSupported()
        {
            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            var data = new object[] { "One" };
            Assert.Equal(
                "One",
                _.CALL(
                    data,
                    new string[0],
                    _.ARGS.Val("0")
                )
            );
        }

        [Fact]
        public void ByRefArgumentIsUpdatedAfterCall()
        {
            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            object arg0 = 1;
            _.CALL(this, "ByRefArgUpdatingFunction", _.ARGS.Ref(arg0, v => { arg0 = v; }).Val(false));
            Assert.Equal(123, arg0);
        }

        [Fact]
        public void ByRefArgumentIsUpdatedAfterCallEvenIfExceptionIsThrown()
        {
            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            object arg0 = 1;
            try
            {
                _.CALL(this, "ByRefArgUpdatingFunction", _.ARGS.Ref(arg0, v => { arg0 = v; }).Val(true));
            }
            catch { }
            Assert.Equal(123, arg0);
        }

        [Fact]
        public void SingleArgumentParamsArrayMethodMayBeCalledWithZeroValues()
        {
            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            Assert.Equal(0, _.CALL(this, "GetNumberOfArgumentsPassedInParamsObjectArray", _.ARGS));
        }

        [Fact]
        public void SingleArgumentParamsArrayMethodMayBeCalledWithSingleValue()
        {
            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            Assert.Equal(1, _.CALL(this, "GetNumberOfArgumentsPassedInParamsObjectArray", _.ARGS.Val(1)));
        }

        [Fact]
        public void SingleArgumentParamsArrayMethodMayBeCalledWithTwoValues()
        {
            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            Assert.Equal(2, _.CALL(this, "GetNumberOfArgumentsPassedInParamsObjectArray", _.ARGS.Val(1).Val(2)));
        }

        public void ByRefArgUpdatingFunction(ref object arg0, bool throwExceptionAfterUpdatingArgument)
        {
            arg0 = 123;
            if (throwExceptionAfterUpdatingArgument)
                throw new Exception("Example exception");
        }

        public Nullable<int> GetNumberOfArgumentsPassedInParamsObjectArray(params object[] args)
        {
            return (args == null) ? (int?)null : args.Length;
        }

        /// <summary>
        /// Then a CALL execution is generated by the translator, the string member accessors should not be renamed - to avoid C# keywords, for example.
        /// If the target is a translated class then any members that use reserved C# keywords must be renamed, but if the target is not a translated
        /// class (a COM component, for example) then the call will fail, so the renaming must not be done at translation time. This means that the
        /// CALL implementation must support name rewriting for when the target is a translated-from-VBScript C# class.
        /// </summary>
        [Fact]
        public void StringMemberAccessorValuesShouldNotBeRewrittenAtTranslationTime()
        {
            var _ = DefaultRuntimeSupportClassFactory.DefaultVBScriptValueRetriever;
            Assert.Equal(
                "Success!",
                _.CALL(new ImpressionOfTranslatedClassWithRewrittenPropertyName(), "Params")
            );
        }

        private class ImpressionOfTranslatedClassWithRewrittenPropertyName
        {
            /// <summary>
            /// This is a property that would have been rewritten from one named "Param" in VBScript (which is valid in VBScript but is a reserved
            /// keyword in C# and so may not appear in C# without some manipulation)
            /// </summary>
            public object rewritten_params { get { return "Success!"; } }
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
    }
}
