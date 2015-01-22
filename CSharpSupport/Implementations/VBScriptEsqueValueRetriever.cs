using CSharpSupport.Attributes;
using CSharpSupport.Exceptions;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace CSharpSupport.Implementations
{
    public class VBScriptEsqueValueRetriever : IAccessValuesUsingVBScriptRules
    {
        // It's feasible that access to the invoker cacher will need to support multi-threaded access depending upon the application in question, so the
        // ConcurrentDictionary seems like a good choice. The most common cases I'm envisaging are for the cache to be built up as various execution paths
        // are followed and then for it to essentially become full. The ConcurrentDictionary allows lock-free reading and so doesn't introduce any costs
        // once this situation is reached.
        private readonly Func<string, string> _nameRewriter;
        private readonly ConcurrentDictionary<InvokerCacheKey, GetInvoker> _getInvokerCache;
        private readonly ConcurrentDictionary<InvokerCacheKey, SetInvoker> _setInvokerCache;
        public VBScriptEsqueValueRetriever(Func<string, string> nameRewriter)
        {
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            _nameRewriter = nameRewriter;
            _getInvokerCache = new ConcurrentDictionary<InvokerCacheKey, GetInvoker>();
            _setInvokerCache = new ConcurrentDictionary<InvokerCacheKey, SetInvoker>();
        }

        /// <summary>
        /// Reduce a reference down to a value type, applying VBScript defaults logic - thrown an exception if this is not possible (null is
        /// acceptable as an input and corresponding return value)
        /// </summary>
        public object VAL(object o)
        {
            if (IsVBScriptValueType(o))
                return o;

            if (IsVBScriptNothing(o))
                throw new ObjectVariableNotSetException();

            var defaultValueFromObject = InvokeGetter(o, null, new object[0]);
            if (IsVBScriptValueType(defaultValueFromObject))
                return defaultValueFromObject;

            // We don't recursively try defaults, so if this default is still not a value type then we're out of luck
            throw new ObjectVariableNotSetException("Object expected (default method/property of object also returned non-value type data)");
        }

        /// <summary>
        /// The comparison (o == VBScriptConstants.Nothing) will return false even if o is VBScriptConstants.Nothing due to the implementation details of
        /// DispatchWrapper. This method delivers a reliable way to test for it.
        /// </summary>
        protected bool IsVBScriptNothing(object o)
        {
            return ((o is DispatchWrapper) && ((DispatchWrapper)o).WrappedObject == null);
        }

        /// <summary>
        /// This will only return a non-VBScript-value-type, if unable to then an exception will be raised (this is used to wrap the right-hand
        /// side of a SET assignment)
        /// </summary>
        public object OBJ(object o)
        {
            if ((o == null) || IsVBScriptValueType(o))
                throw new Exception("Object expected (SET-assignment-derived statement requires a non-value-type)");

            return o;
        }

        /// <summary>
        /// Reduce a reference down to a numeric value type, applying VBScript defaults logic and then trying to parse as a number - throwing
        /// an exception if this is not possible. Null (aka VBScript Empty) is acceptable and will result in zero being returned. DBNull.Value
        /// (aka VBScript Null) is not acceptable and will result in an exception being raised, as any other invalid value (eg. a string or
        /// an object without an appropriate default property member) will. This is used by translated code and is very similar to the
        /// IProvideVBScriptCompatFunctionalityToIndividualRequests.CDBL method, though it may apply different rules where appropriate
        /// since it is not bound to the behaviour of a built-in VBScript method - for example, if the input is a date, then that
        /// date will be returned unaltered
        /// </summary>
        public double NUM(object o)
        {
            // Get the null-esque cases out of the way
            if (o == null)
                return 0;
            if (o == DBNull.Value)
                throw new ArgumentException("Invalid use of Null");

            // Now some other quick and easy (and common) ones (this method is used by the FOR translation so having this quick casts is helpful)
            if (o is bool)
            {
                if ((bool)o)
                    return -1;
                return 0;
            }
            if (o is double)
                return (double)o;
            if (o is int)
                return (double)(int)o;

            // Try to extract the value-type data from the reference, dealing with the false-Empty/Null cases first
            var value = VAL(o);
            if (value == null)
                return 0;
            if (value == DBNull.Value)
                throw new Exception("Invalid use of Null");

            // Then fall back to parsing as a string
            var valueString = value.ToString();
            double parsedValue;
            if (!double.TryParse(valueString, out parsedValue))
                throw new TypeMismatchException("[string \"" + valueString + "\"] (unable to translate into a numeric values)");
            return parsedValue;
        }

        /// <summary>
        /// This wraps a call to NUM and allows an exception to be made for DBNull.Value (VBScript Null) in that the same value will be returned
        /// (it is not a valid input for NUM).
        /// </summary>
        public object NULLABLENUM(object o)
        {
            return (o == DBNull.Value) ? DBNull.Value : (object)NUM(o);
        }

        /// <summary>
        /// Reduce a reference down to a string value type (in most cases), applying VBScript defaults logic and then taking a string representation.
        /// Null (aka VBScript Empty) is acceptable and will result in null being returned. DBNull.Value (aka VBScript Null) is also acceptable and
        /// will also result in itself being returned - this is the only case in which a non-null-and-non-string value will be returned. This
        /// conversion should only used for comparisons with string literals, where special rules apply (which makes the method slightly
        /// less useful than NUM, which is used in comparisons with numeric literals but also in some other cases, such as FOR loops).
        /// </summary>
        public object STR(object o)
        {
            // Get the null-esque cases out of the way
            if ((o == null) || (o == DBNull.Value))
                return o;

            // Try to extract the value-type data from the reference, dealing with the same null cases as above first
            var value = VAL(o);
            if ((value == null) || (value == DBNull.Value))
                return value;

            // Then we only need to call ToString (this, thankfully, works correctly for booleans - the casing is consistent between C# and VBScript)
            return value.ToString();
        }

        /// <summary>
        /// Layer an enumerable wrapper over a reference, if possible (an exception will be thrown if not)
        /// </summary>
        public IEnumerable ENUMERABLE(object o)
        {
            if (o == null)
                throw new ArgumentNullException("o");

            // VBScript will only consider object references to be enumerable (unlike C#, which will consider a string to be an enumerable set
            // characters, for example)
            if (IsVBScriptValueType(o))
                throw new ArgumentException("Object not a collection");

            // Try casting to IEnumerable first - it's the easiest approach and will work with (many) managed references and some COM object
            var enumerable = o as IEnumerable;
            if (enumerable != null)
                return enumerable;

            // Failing that, attempt access through IDispatch - try calling the method with DispId -4 (if there is one) and casting the return
            // value to IEnumVariant and then wrapping up into a managed IEnumerable
            if (IDispatchAccess.ImplementsIDispatch(o))
            {
                object enumerator;
                try
                {
                    enumerator = IDispatchAccess.Invoke<object>(o, IDispatchAccess.InvokeFlags.DISPATCH_METHOD, -4);
                }
                catch (MissingMemberException)
                {
                    throw new ArgumentException("IDispatch reference does not have a method with DispId -4");
                }
                var enumeratorAsEnumVariant = enumerator as IEnumVariant;
                if (enumeratorAsEnumVariant == null)
                    throw new ArgumentException("IDispatch reference has a DispId -4 return value that does not implement IEnumVariant");
                return new ManagedEnumeratorWrapper(new IDispatchEnumeratorWrapper(enumeratorAsEnumVariant));
            }

            // Give up and throw the VBScript error message
            throw new ArgumentException("Object not a collection");
        }

        /// <summary>
        /// Reduce a reference down to a boolean, throwing an exception if this is not possible. This will apply the same logic as VAL but then
        /// require a numeric value or null, otherwise an exception will be raised. Zero and null equate to false, non-zero numbers to true.
        /// </summary>
        public bool IF(object o)
        {
            if (o == null)
                return false;

            // Try to extract the value-type data from the reference, dealing with the false-Empty/Null cases first
            var value = VAL(o);
            if ((value == null) || (value == DBNull.Value))
                return false;

            // Then fall back to parsing as a string
            var valueString = value.ToString();
            double parsedValue;
            if (!double.TryParse(valueString, out parsedValue))
                throw new TypeMismatchException("[string \"" + valueString + "\"] (unable to translate into boolean for IF statement)");
            return parsedValue != 0;
        }

        /// <summary>
        /// This is used to wrap arguments such that those that must be passed ByVal can have changes reflected after a method call completes
        /// </summary>
        public IBuildCallArgumentProviders ARGS
        {
            get { return new DefaultCallArgumentProvider(this); }
        }

        /// <summary>
        /// This requires a target with optional member accessors and arguments - eg. "Test" is a target only, "a.Test" has target "a" with one
        /// named member "Test", "a.Test(0)" has target "a", named member "Test" and a single argument "0". The expression "a(Test(0))" would
        /// require nested CALL executions, one with target "Test" and a single argument "0" and a second with target "a" and a single
        /// argument which was the result of the first call.
        /// </summary>
        public object CALL(object target, IEnumerable<string> members, IProvideCallArguments argumentProvider)
        {
            if (members == null)
                throw new ArgumentNullException("members");
            if (argumentProvider == null)
                throw new ArgumentNullException("argumentProvider");

            var arguments = argumentProvider.GetInitialValues().ToArray();
            var returnValue = CALL(target, members, arguments);
            for (var index = 0; index < arguments.Length; index++)
                argumentProvider.OverwriteValueIfByRef(index, arguments[index]);
            return returnValue;
        }

        /// <summary>
        /// Note: The arguments array elements may be mutated if the call target has "ref" method arguments.
        /// </summary>
        private object CALL(object target, IEnumerable<string> members, object[] arguments)
        {
            if (members == null)
                throw new ArgumentNullException("members");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            var memberAccessorsArray = members.ToArray();
            if (memberAccessorsArray.Any(m => string.IsNullOrWhiteSpace(m)))
                throw new ArgumentException("Null/blank value in members set");

            // If there are no member accessor or arguments then just return the object directly (if the caller wants this to be a value type
            // then they'll have to use VAL to try to apply that requirement). This is the only case in which it's acceptable for target to
            // be null.
            if (!memberAccessorsArray.Any() && !arguments.Any())
                return target;
            else if (target == null)
                throw new ArgumentException("The target reference may only be null if there are no member accessors or arguments specified");

            // If there are no member accessors but there are arguments then
            // 1. Attempt array access (this is the only time that array access is acceptable (o.Names(0) is not allowed by VBScript if the
            //    "Names" property is an array, it will only work if Names if a Function or indexed Property)
            // 2. If target implements IDispatch then try accessing the DispId zero method or property, passing the arguments
            // 3. If it's not an IDispatch reference, then the arguments will be passed to a method or indexed property that will accept them,
            //    taking into account the IsDefault attribute
            if (!memberAccessorsArray.Any() && arguments.Any())
                return InvokeGetter(target, null, arguments);

            // If there are member accessors but no arguments then we walk down each member accessor, no defaults are considered
            if (!arguments.Any())
                return WalkMemberAccessors(target, memberAccessorsArray);

            // If there member accessors AND arguments then all-but-the-last member accessors should be walked through as argument-less lookups
            // and the final member accessor should be a method or property call whose name matches the member accessor. Note that the arguments
            // can never be for an array look up if there are member accessors, it must always be a function or property.
            target = WalkMemberAccessors(target, memberAccessorsArray.Take(memberAccessorsArray.Length - 1));
            var finalMemberAccessor = memberAccessorsArray[memberAccessorsArray.Length - 1];
            if (target == null)
                throw new ArgumentException("Unable to access member \"" + finalMemberAccessor + "\" on null reference");
            return InvokeGetter(target, finalMemberAccessor, arguments);
        }

        /// <summary>
        /// This will throw an exception for null target or arguments references or if the setting fails (eg. invalid number of arguments,
        /// invalid member accessor - if specified - argument thrown by the target setter). This must not be called with a target reference
        /// only (null optionalMemberAccessor and zero arguments) as it would need to change the caller's reference to target, which is not
        /// possible (in that case, a straight assignment should be generated - no call to SET required). Note that the valueToSetTo argument
        /// comes before any others since VBScript will evaulate the right-hand side of the assignment before the left, which may be important
        /// if an error is raised at some point in the operation.
        /// </summary>
        public void SET(object valueToSetTo, object target, string optionalMemberAccessor, IProvideCallArguments argumentProvider)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (argumentProvider == null)
                throw new ArgumentNullException("argumentProvider");

            var arguments = argumentProvider.GetInitialValues().ToArray();
            if ((optionalMemberAccessor == null) && !arguments.Any())
                throw new ArgumentException("This must be called with a non-null optionalMemberAccessor and/or one or more arguments, null optionalMemberAccessor and zero arguments is not supported");

            var cacheKey = new InvokerCacheKey(target.GetType(), optionalMemberAccessor, arguments.Length);
            SetInvoker invoker;
            if (!_setInvokerCache.TryGetValue(cacheKey, out invoker))
            {
                invoker = GenerateSetInvoker(target, optionalMemberAccessor, arguments);
                _setInvokerCache.TryAdd(cacheKey, invoker);
            }
            invoker(target, arguments, valueToSetTo);
            for (var index = 0; index < arguments.Length; index++)
                argumentProvider.OverwriteValueIfByRef(index, arguments[index]);
        }

        /// <summary>
        /// The arguments set must be an array since its contents may be mutated if the call target has "ref" parameters
        /// </summary>
        private object InvokeGetter(object target, string optionalName, object[] arguments)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            var cacheKey = new InvokerCacheKey(target.GetType(), optionalName, arguments.Length);
            GetInvoker invoker;
            if (!_getInvokerCache.TryGetValue(cacheKey, out invoker))
            {
                invoker = GenerateGetInvoker(target, optionalName, arguments);
                _getInvokerCache.TryAdd(cacheKey, invoker);
            }
            return invoker(target, arguments);
        }

        private delegate object GetInvoker(object target, object[] arguments);
        private GetInvoker GenerateGetInvoker(object target, string optionalName, IEnumerable<object> arguments)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            var argumentsArray = arguments.ToArray();
            var targetType = target.GetType();
            if (targetType.IsArray)
            {
                if (targetType.GetArrayRank() != argumentsArray.Length)
                    throw new ArgumentException("Argument count (" + argumentsArray.Length + ") does not match arrary rank (" + targetType.GetArrayRank() + ")");

                var arrayTargetParameter = Expression.Parameter(typeof(object), "target");
                var indexesParameter = Expression.Parameter(typeof(object[]), "arguments");
                var arrayAccessExceptionParameter = Expression.Parameter(typeof(Exception), "e");
                return Expression.Lambda<GetInvoker>(
                    Expression.TryCatch(
                        Expression.Convert(
                            Expression.ArrayAccess(
                                Expression.Convert(arrayTargetParameter, targetType),
                                Enumerable.Range(0, argumentsArray.Length).Select(index =>
                                    GetVBScriptStyleArrayIndexParsingExpression(
								        Expression.ArrayAccess(indexesParameter, Expression.Constant(index))
                                    )
                                )
                            ),
                            typeof(object) // Without this we may get an "Expression of type 'System.Int32' cannot be used for return type 'System.Object'" or similar
                        ),
                        Expression.Catch(
                            arrayAccessExceptionParameter,
                            Expression.Throw(
                                GetNewArgumentException(
                                    "Error accessing array with specified indexes (likely a non-numeric array index/argument or index-out-of-bounds)",
                                    arrayAccessExceptionParameter
                                ),
                                typeof(object)
                            )
                        )
                    ),
                    new[]
			        {
				        arrayTargetParameter,
				        indexesParameter
			        }
                ).Compile();
            }

            var errorMessageMemberDescription = (optionalName == null) ? "default member" : ("member \"" + optionalName + "\"");
            if (argumentsArray.Length == 0)
                errorMessageMemberDescription = "parameter-less " + errorMessageMemberDescription;
            else
                errorMessageMemberDescription += " that will accept " + argumentsArray.Length + " argument(s)";

            if (IDispatchAccess.ImplementsIDispatch(target))
            {
                int dispId;
                if (optionalName == null)
                    dispId = 0;
                else
                {
                    try
                    {
                        // We don't use the nameRewriter here since we won't have rewritten the COM component, it's the C# generated from the
                        // VBScript source that we may have rewritten
                        dispId = IDispatchAccess.GetDispId(target, optionalName);
                    }
                    catch (Exception e)
                    {
                        throw new ObjectVariableNotSetException("Unable to identify " + errorMessageMemberDescription + " (target implements IDispatch)", e);
                    }
                }
                return (invokeTarget, invokeArguments) =>
                {
                    try
                    {
                        return IDispatchAccess.Invoke<object>(
                            invokeTarget,
                            IDispatchAccess.InvokeFlags.DISPATCH_METHOD | IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYGET,
                            dispId,
                            invokeArguments.ToArray()
                        );
                    }
                    catch (Exception e)
                    {
                        throw new ArgumentException("Error executing " + errorMessageMemberDescription + " (target implements IDispatch): " + e.GetBaseException(), e);
                    }
                };
            }

            var possibleMethods = (optionalName == null)
                ? GetDefaultGetMethods(targetType, argumentsArray.Length)
                : GetNamedGetMethods(targetType, optionalName, argumentsArray.Length);
            if (!possibleMethods.Any())
                throw new ObjectVariableNotSetException("Unable to identify " + errorMessageMemberDescription);

            var method = possibleMethods.First();
            var targetParameter = Expression.Parameter(typeof(object), "target");
            var argumentsParameter = Expression.Parameter(typeof(object[]), "arguments");
            var exceptionParameter = Expression.Parameter(typeof(Exception), "e");

            // If there are any ByRef arguments then we need local variables that are of ByRef types. These will be populated with values from the
            // arguments array before the method call, then used AS arguments in the method call, then their values copied back into the slots in
            // the arguments array that they came from.
            var methodParameters = method.GetParameters();
            var argParameters = methodParameters.Select((p, index) => new { Parameter = p, Index = index });
            var argExpressions = argParameters.Select(a =>
                a.Parameter.ParameterType.IsByRef
                    ? (Expression)Expression.Variable(
                        GetNonByRefType(a.Parameter.ParameterType),
                        a.Parameter.Name
                    )
                    : Expression.Convert(
                        Expression.ArrayAccess(argumentsParameter, Expression.Constant(a.Index)),
                        a.Parameter.ParameterType
                    )
                )
                .ToArray();
            var byRefArgAssignmentsForMethodCall = argExpressions
                .Select((a, index) => new { Parameter = a, Index = index })
                .Where(a => a.Parameter is ParameterExpression)
                .Select(a =>
                    Expression.Assign(
                        a.Parameter,
                        Expression.Convert(
                            Expression.ArrayAccess(argumentsParameter, Expression.Constant(a.Index)),
                            a.Parameter.Type
                        )
                    )
                );
            var byRefArgAssignmentsForReturn = argExpressions
                .Select((a, index) => new { Parameter = a, Index = index })
                .Where(a => a.Parameter is ParameterExpression)
                .Select(a =>
                    Expression.Assign(
                        Expression.ArrayAccess(argumentsParameter, Expression.Constant(a.Index)),
                        a.Parameter
                    )
                );

            var methodCall = Expression.Call(
                Expression.Convert(targetParameter, targetType),
                method,
                argExpressions
            );

            var resultVariable = Expression.Variable(typeof(object));
            Expression[] methodCallAndAndResultAssignments;
            if (method.ReturnType == typeof(void))
            {
                // If the method has no return type then we'll need to just return null since the GetInvoker delegate that
                // we're trying to compile to has a return type of "object" (so execute methodCall and then set the result
                // variable to null)
                methodCallAndAndResultAssignments = new Expression[]
                {
                    methodCall,
                    Expression.Assign(
                        resultVariable,
                        Expression.Constant(null, typeof(object))
                    )
                };
            }
            else
            {
                // If there IS a a return type then assign the method call's return value to the result variable
                methodCallAndAndResultAssignments = new Expression[]
                {
                    Expression.Assign(
                        resultVariable,
                        methodCall
                    )
                };
            }

            // The Throw expression requires a return type to be specified since the Try block has a return type - without
            // this a runtime "Body of catch must have the same type as body of try" exception will be raised
            return Expression.Lambda<GetInvoker>(
                Expression.TryCatch(
                    Expression.Block(
                        argExpressions
                            .Where(a => a is ParameterExpression)
                            .Cast<ParameterExpression>()
                            .Concat(new[] { resultVariable })
                            .ToArray(),
                        byRefArgAssignmentsForMethodCall
                            .Concat(
                                methodCallAndAndResultAssignments
                            )
                            .Concat(byRefArgAssignmentsForReturn)
                            .Concat(
                                new[] { resultVariable }
                            )
                            .ToArray()
                    ),
                    Expression.Catch(
                        exceptionParameter,
                        Expression.Throw(
                            GetNewArgumentException("Error executing " + errorMessageMemberDescription, exceptionParameter),
                            typeof(object)
                        )
                    )
                ),
                targetParameter,
                argumentsParameter
            ).Compile();
        }

        private delegate void SetInvoker(object target, object[] arguments, object value);
        private SetInvoker GenerateSetInvoker(object target, string optionalMemberAccessor, IEnumerable<object> arguments)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            // If there are no member accessors but there are arguments then firstly attempt array access, this is the only time that array access is acceptable
            // (o.Names(0) is not allowed by VBScript if the "Names" property is an array, it will only work if Names is a Function or indexed Property and since
            // we're attempting a set here, it will only work if it's an indexed Property)
            var argumentsArray = arguments.ToArray();
            var targetType = target.GetType();
            if ((optionalMemberAccessor == null) && targetType.IsArray)
            {
                if (targetType.GetArrayRank() != argumentsArray.Length)
                    throw new ArgumentException("Argument count (" + argumentsArray.Length + ") does not match arrary rank (" + targetType.GetArrayRank() + ")");

                // Without the targetType.GetElementType() specified for the Expression.Catch, a "Body of catch must have the same type as body of try" exception
                // will be raised. Since I would have expected the try block to return nothing (void) I'm not quite sure why this is.
                var arrayTargetParameter = Expression.Parameter(typeof(object), "target");
                var indexesParameter = Expression.Parameter(typeof(object[]), "arguments");
                var arrayAccessExceptionParameter = Expression.Parameter(typeof(Exception), "e");
                var arrayValueParameter = Expression.Parameter(typeof(object), "value");
                return Expression.Lambda<SetInvoker>(
                    Expression.TryCatch(
                        Expression.Assign(
                            Expression.ArrayAccess(
                                Expression.Convert(arrayTargetParameter, targetType),
                                Enumerable.Range(0, argumentsArray.Length).Select(index =>
                                    GetVBScriptStyleArrayIndexParsingExpression(
                                        Expression.ArrayAccess(indexesParameter, Expression.Constant(index))
                                    )
                                )
                            ),
                            Expression.Convert(
                                arrayValueParameter,
                                targetType.GetElementType()
                            )
                        ),
                        Expression.Catch(
                            arrayAccessExceptionParameter,
                            Expression.Throw(
                                GetNewArgumentException(
                                    "Error accessing array with specified indexes (likely a non-numeric array index/argument or index-out-of-bounds)",
                                    arrayAccessExceptionParameter
                                ),
                                targetType.GetElementType()
                            )
                        )
                    ),
                    new[]
			        {
				        arrayTargetParameter,
				        indexesParameter,
                        arrayValueParameter
			        }
                ).Compile();
            }

            var errorMessageMemberDescription = (optionalMemberAccessor == null) ? "default member" : ("member \"" + optionalMemberAccessor + "\"");
            if (argumentsArray.Length == 0)
                errorMessageMemberDescription = "parameter-less " + errorMessageMemberDescription;
            else
                errorMessageMemberDescription += " that will accept " + argumentsArray.Length + " argument(s)";

            if (IDispatchAccess.ImplementsIDispatch(target))
            {
                int dispId;
                if (optionalMemberAccessor == null)
                    dispId = 0;
                else
                {
                    try
                    {
                        // We don't use the nameRewriter here since we won't have rewritten the COM component, it's the C# generated from the
                        // VBScript source that we may have rewritten
                        dispId = IDispatchAccess.GetDispId(target, optionalMemberAccessor);
                    }
                    catch (Exception e)
                    {
                        throw new ArgumentException("Unable to identify " + errorMessageMemberDescription + " (target implements IDispatch)", e);
                    }
                }
                return (invokeTarget, invokeArguments, value) =>
                {
                    try
                    {
                        IDispatchAccess.Invoke<object>(
                            target,
                            IDispatchAccess.InvokeFlags.DISPATCH_PROPERTYPUT,
                            dispId,
                            argumentsArray.Concat(new[] { value }).ToArray()
                        );
                    }
                    catch (Exception e)
                    {
                        throw new ArgumentException("Error executing " + errorMessageMemberDescription + " (target implements IDispatch): " + e.GetBaseException(), e);
                    }
                };
            }

            MethodInfo method;
            if (optionalMemberAccessor != null)
            {
                // If there is a non-null optionalMemberAccessor but no arguments then try setting the non-indexed property, no defaults considered.
                // If there is a non-null optionalMemberAccessor and there are arguments then no defaults are considered and the member accessor
                // must be an indexed property, it can not be an array (see note above about the only place that array access is permitted).
                method = GetNamedSetMethods(targetType, optionalMemberAccessor, argumentsArray.Length).FirstOrDefault();
            }
            else
            {
                // Try accessing a default indexed property (either a a native C# property or a method with IsDefault and TranslatedProperty attributes)
                method = GetDefaultSetMethods(targetType, argumentsArray.Length).FirstOrDefault();
            }
            if (method == null)
                throw new ArgumentException("Unable to identify " + errorMessageMemberDescription);

            var targetParameter = Expression.Parameter(typeof(object), "target");
            var argumentsParameter = Expression.Parameter(typeof(object[]), "arguments");
            var exceptionParameter = Expression.Parameter(typeof(Exception), "e");
            var valueParameter = Expression.Parameter(typeof(object), "value");
            return Expression.Lambda<SetInvoker>(

                Expression.TryCatch(

                    Expression.Call(
                        Expression.Convert(targetParameter, targetType),
                        method,
                        method.GetParameters().Select((arg, index) =>
                        {
                            // The method will have one more parameter than there are arguments since the last parameter is the value to set to
                            return Expression.Convert(
                                (index == argumentsArray.Length)
                                    ? (Expression)valueParameter
                                    : Expression.ArrayAccess(argumentsParameter, Expression.Constant(index)),
                                arg.ParameterType
                            );
                        })
                    ),

                    Expression.Catch(
                        exceptionParameter,
                        Expression.Throw(
                            GetNewArgumentException("Error executing " + errorMessageMemberDescription, exceptionParameter)
                        )
                    )

                ),
                new[]
				{
					targetParameter,
					argumentsParameter,
                    valueParameter
				}
            ).Compile();
        }

        /// <summary>
        /// VBScript expects integer array index values but will accept fractional values or even strings representing integers or fractional values.
        /// Any fractional values are rounded to the closest even number (so 1.5 rounds to 1, 2.5 also rounds to 2, 3.5 rounds to 4, etc..)
        /// </summary>
        private Expression GetVBScriptStyleArrayIndexParsingExpression(Expression index)
        {
            if (index == null)
                throw new ArgumentNullException("index");

            var returnValueParameter = Expression.Parameter(typeof(int), "retVal");
            return Expression.Block(
                typeof(int),
                new[] { returnValueParameter },
                Expression.IfThenElse(
                    Expression.TypeIs(index, typeof(int)),
                    Expression.Assign(
                        returnValueParameter,
                        Expression.Convert(
                            index,
                            typeof(int)
                        )
                    ),
                    Expression.IfThenElse(
                        Expression.TypeIs(index, typeof(double)),
                        Expression.Assign(
                            returnValueParameter,
                            Expression.Convert(
                                Expression.Call(
                                    null,
                                    typeof(Math).GetMethod("Round", new[] { typeof(double), typeof(MidpointRounding) }),
                                    Expression.Convert(
                                        index,
                                        typeof(double)
                                    ),
                                    Expression.Constant(MidpointRounding.ToEven)
                                ),
                                typeof(int)
                            )
                        ),
                        Expression.Assign(
                            returnValueParameter,
                            Expression.Convert(
                                Expression.Call(
                                    null,
                                    typeof(Math).GetMethod("Round", new[] { typeof(double), typeof(MidpointRounding) }),
                                    Expression.Call(
                                        null,
                                        typeof(double).GetMethod("Parse", new[] { typeof(string) }),
                                        Expression.Call(
                                            index,
                                            typeof(object).GetMethod("ToString")
                                        )
                                    ),
                                    Expression.Constant(MidpointRounding.ToEven)
                                ),
                                typeof(int)
                            )
                        )
                    )
                ),
                returnValueParameter
            );
        }

        private Expression GetNewArgumentException(string message, ParameterExpression exceptionParameter)
        {
            if (string.IsNullOrWhiteSpace(message))
                throw new ArgumentException("Null/blank message specified");
            if (exceptionParameter == null)
                throw new ArgumentNullException("exceptionParameter");

            return Expression.New(
                typeof(ArgumentException).GetConstructor(new[] { typeof(string), typeof(Exception) }),
                Expression.Call(
                    typeof(string).GetMethod("Concat", new[] { typeof(object), typeof(object) }),
                    Expression.Constant(message.Trim() + ": "),
                    Expression.Call(
                        exceptionParameter,
                        typeof(Exception).GetMethod("GetBaseException")
                    )
                ),
                exceptionParameter
            );
        }

        private Type GetNonByRefType(Type byRefType)
        {
            if (byRefType == null)
                throw new ArgumentNullException("byRefType");
            if (!byRefType.IsByRef)
                throw new ArgumentException("Must be a type which reports IsByRef true");

            var byRefFullName = byRefType.FullName;
            var nonByRefFullName = byRefFullName.Substring(0, byRefFullName.Length - 1);
            return Type.GetType(
                byRefType.AssemblyQualifiedName.Replace(byRefFullName, nonByRefFullName),
                true
            );
        }

        private int ApplyVBScriptIndexLogicToNonIntegerValue(double value)
        {
            return (int)Math.Round(value, MidpointRounding.ToEven); // This is what effectively what VBScript does
        }

        private object WalkMemberAccessors(object target, IEnumerable<string> memberAccessors)
        {
            if (target == null)
                throw new ArgumentNullException("target");
            if (memberAccessors == null)
                throw new ArgumentNullException("memberAccessors");

            foreach (var memberAccessor in memberAccessors)
            {
                if (string.IsNullOrWhiteSpace(memberAccessor))
                    throw new ArgumentException("Encountered null/blank in memberAccessors set");

                // This should only be the case after we've gone round the loop at least once
                if (target == null)
                    throw new ArgumentException("Unable to access member \"" + memberAccessor + "\" on null reference");

                target = InvokeGetter(target, memberAccessor, new object[0]);
            }
            return target;
        }

        private IEnumerable<MethodInfo> GetDefaultGetMethods(Type type, int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            return GetGetMethods(type, null, DefaultMemberBehaviourOptions.MustBeDefault, MemberNameMatchBehaviourOptions.Precise, numberOfArguments);
        }

        private IEnumerable<MethodInfo> GetDefaultSetMethods(Type type, int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            return GetSetMethods(type, null, DefaultMemberBehaviourOptions.MustBeDefault, MemberNameMatchBehaviourOptions.Precise, numberOfArguments);
        }

        private IEnumerable<MethodInfo> GetNamedGetMethods(Type type, string name, int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Null/blank name specified");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            // There the nameRewriter WILL be considered in case it's trying to access classes we've translated. However, there's also a chance that
            // we could be accessing a non-IDispatch CLR type from somewhere, so GetNamedGetMethods will try to match using the nameRewriter first and
            // then fallback to a perfect match non-rewritten name and finally to a case-insensitive match to a non-rewritten name.
            return GetGetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.UseNameRewriter, numberOfArguments)
                .Concat(
                    GetGetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.Precise, numberOfArguments)
                )
                .Concat(
                    GetGetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.CaseInsensitive, numberOfArguments)
                );
        }

        private IEnumerable<MethodInfo> GetNamedSetMethods(Type type, string name, int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Null/blank name specified");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            // There the nameRewriter WILL be considered in case it's trying to access classes we've translated. However, there's also a chance that
            // we could be accessing a non-IDispatch CLR type from somewhere, so GetNamedGetMethods will try to match using the nameRewriter first and
            // then fallback to a perfect match non-rewritten name and finally to a case-insensitive match to a non-rewritten name.
            return GetSetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.UseNameRewriter, numberOfArguments)
                .Concat(
                    GetSetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.Precise, numberOfArguments)
                )
                .Concat(
                    GetSetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.CaseInsensitive, numberOfArguments)
                );
        }

        private IEnumerable<MethodInfo> GetGetMethods(
            Type type,
            string optionalName,
            DefaultMemberBehaviourOptions defaultMemberBehaviour,
            MemberNameMatchBehaviourOptions memberNameMatchBehaviour,
            int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (!Enum.IsDefined(typeof(DefaultMemberBehaviourOptions), defaultMemberBehaviour))
                throw new ArgumentOutOfRangeException("defaultMemberBehaviour");
            if (!Enum.IsDefined(typeof(MemberNameMatchBehaviourOptions), memberNameMatchBehaviour))
                throw new ArgumentOutOfRangeException("memberNameMatchBehaviour");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            // There is some crazy behaviour around default member accesses when VBScript and COM are combined, which need to be reflected here. For
            // example, if C# class is wrapped up as a COM object and consumed by VBScript and that class is decorated with ComVisible(true) and an
            // unnamed parameter-less member access is made against that object - eg.
            //
            //   Set o = CreateObject("My.ComComponent")
            //   If (o = "Test") Then
            //    ' Something
            //   End If
            // 
            // then various options are considered. If there is a method or property with DispId(0) then that is an obvious place to start, but an
            // error will be raised if has arguments, since zero arguments are being passed in. Also, if multiple members are marked with DispId(0)
            // then it doesn't know which to choose and they are all ignored. In this case, or the case of no DispId(0) members, there are further
            // alternatives considered: If there is a method or property called "Value" with zero arguments then this will queried. If this is not
            // available then the parameter-less "ToString" method will be called - but only if the type or one of its explicit base types declares
            // ComVisible(true). All classes implicitly inherit the base type Object, which declares ComVisible(true), but if that is the only type
            // in the inheritance tree that does then the "ToString" fallback is not considered. The "Value" property / method fallback is acceptable
            // if it is on a type with ComVisible(false) so long as elsewhere in the inheritance tree there is a type other than Object which specifies
            // ComVisible(true). This magic is actually something that happens when the class is exposed as COM and introduced by the C# compiler, it
            // will set the ToString method to have DispId zero if nothing else does and if DispId(0) attributes are found for multiple methods (any-
            // where in a type's inheritance tree) then it will also force the ToString method to have DispId zero (this is talked about in the book
            // ".NET and COM: The Complete Interoperability Guide", though I'm yet to find a definitive reference for the default "Value" member
            // assignment). Depending upon how the component is exposed when consumed through .net, none of this may be required - if we get true
            // back when we call IDispatchAccess.ImplementsIDispatch(target) then IDispatch will be used to query the object and this code path
            // will not be visited. This reflection-based approach is much more complex, but it allows us to cache a lambda for this particular
            // member access on this target type, which means that we'll get some performance benefits if it's called many times. I fully suspect
            // that, at this time, there are various scenarios which are not being handled correctly - I'm trying to cover the obvious cases that
            // VBScript code would have to deal with and need to read ".NET and COM: The Complete Interoperability Guide" more thoroughly and try
            // to apply all of that knowledge when I absorb it!
            //
            // One final thing to consider: these special cases should not apply to classes that were translated from VBScript source since these edge
            // cases are something applied by the C# compiler and so would not have been applied to the original VBScript classes (if a VBScript class
            // has no default method or property and is accessed in a manner as illustrated above, where the object reference needs to be processed as
            // a value reference, then a Type Mismatch error would be raised - a parameter-less "Value" property would be no help if it wasn't explicitly
            // declared to be default). Translated classes will have the "SourceClassName" attribute specifiede on them, which is what the following code
            // uses to differentiate between translated classes and everything else.

            // Try to locate methods and/or properties that meet the criteria; have the correct number of arguments and match the name or are default, if
            // default member access is desired (dealing with the differing logic for translated-from-VBScript types and non-translated-from-VBScript).
            var nameMatcher = (optionalName != null) ? GetNameMatcher(optionalName, memberNameMatchBehaviour) : (name => true);
            var typeWasTranslatedFromVBScript = TypeIsTranslatedFromVBScript(type);
            var typeIsComVisible = TypeIsComVisible(type);
            var typeHasAnyDispIdZeroMember = AnyDispIdZeroMemberExists(type);
            var typeHasAmbiguousDispIdZeroMember = typeHasAnyDispIdZeroMember && DispIdZeroIsAmbiguous(type);
            var allMethods = GetMethodsThatAreNotRelatedToProperties(type);
            var applicableMethods = allMethods
                    .Where(m => nameMatcher(m.Name))
                    .Where(m => m.GetParameters().Length == numberOfArguments)
                    .Where(m =>
                        (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                        (m.GetCustomAttributes(typeof(IsDefault), true).Any()) ||
                        (!typeWasTranslatedFromVBScript && typeIsComVisible && !typeHasAmbiguousDispIdZeroMember && MemberHasDispIdZero(m))
                    );
            var readableProperties = type.GetProperties().Where(p => p.CanRead);
            var applicableProperties = readableProperties
                .Where(p => nameMatcher(p.Name))
                .Where(p => p.GetIndexParameters().Length == numberOfArguments)
                .Where(p =>
                    (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                    (p.GetCustomAttributes(typeof(IsDefault), true).Any()) ||
                    (!typeWasTranslatedFromVBScript && typeIsComVisible && !typeHasAmbiguousDispIdZeroMember && MemberHasDispIdZero(p))
                )
                .Select(p => p.GetGetMethod());

            // If no matches were found and we're looking for a parameter-less default member access and the target type was not translated from VBScript
            // source, then apply the other fallbacks that the C# compiler would have added to the IDispatch interface
            var allOptions = applicableMethods.Concat(applicableProperties);
            if (!allOptions.Any() && (defaultMemberBehaviour == DefaultMemberBehaviourOptions.MustBeDefault) && (numberOfArguments == 0) && !typeWasTranslatedFromVBScript && typeIsComVisible)
            {
                allOptions = allMethods.Where(m => !m.GetParameters().Any() && m.Name.Equals("Value", StringComparison.OrdinalIgnoreCase))
                    .Concat(readableProperties.Where(p => !p.GetIndexParameters().Any() && p.Name.Equals("Value", StringComparison.OrdinalIgnoreCase)).Select(p => p.GetGetMethod()));
                if (!allOptions.Any())
                    allOptions = new[] { type.GetMethod("ToString", Type.EmptyTypes) };
            }

            // In the cases where multiple options are identified, sort by the most specific (members declared on the current type those declared further
            // down in the inheritance tree)
            return allOptions.OrderByDescending(m => GetMemberInheritanceDepth(m, type));
        }

        private IEnumerable<MethodInfo> GetSetMethods(
            Type type,
            string optionalName,
            DefaultMemberBehaviourOptions defaultMemberBehaviour,
            MemberNameMatchBehaviourOptions memberNameMatchBehaviour,
            int numberOfArguments)
        {
            if (type == null)
                throw new ArgumentNullException("type");
            if (!Enum.IsDefined(typeof(DefaultMemberBehaviourOptions), defaultMemberBehaviour))
                throw new ArgumentOutOfRangeException("defaultMemberBehaviour");
            if (!Enum.IsDefined(typeof(MemberNameMatchBehaviourOptions), memberNameMatchBehaviour))
                throw new ArgumentOutOfRangeException("memberNameMatchBehaviour");
            if (numberOfArguments < 0)
                throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

            var nameMatcher = (optionalName != null) ? GetNameMatcher(optionalName, memberNameMatchBehaviour) : (name => true);
            return
                GetMethodsThatAreNotRelatedToProperties(type)
                    .Where(m => nameMatcher(m.Name))
                    .Where(m => m.GetParameters().Length == (numberOfArguments + 1)) // Method takes property arguments plus one for the value
                    .Where(m => m.GetCustomAttributes(true).Cast<Attribute>().Any(a => a is TranslatedProperty))
                    .Where(m =>
                        (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                        (m.GetCustomAttributes(true).Cast<Attribute>().Any(a => a is IsDefault))
                    )
                .Concat(
                    type.GetProperties()
                        .Where(p => p.CanWrite)
                        .Where(p => nameMatcher(p.Name))
                        .Where(p => p.GetIndexParameters().Length == numberOfArguments)
                        .Where(p =>
                            (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                            (p.GetCustomAttributes(true).Cast<Attribute>().Any(a => a is IsDefault))
                        )
                        .Select(p => p.GetSetMethod())
                );
        }

        private IEnumerable<MethodInfo> GetMethodsThatAreNotRelatedToProperties(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            return type.GetMethods() // This gets all public methods, whether they're declared in the specified type or anywhere in its inheritance tree
                .Except(type.GetProperties().Where(p => p.CanRead).Select(p => p.GetGetMethod()))
                .Except(type.GetProperties().Where(p => p.CanWrite).Select(p => p.GetSetMethod()));
        }

        private int GetMemberInheritanceDepth(MemberInfo memberInfo, Type type)
        {
            if (memberInfo == null)
                throw new ArgumentNullException("memberInfo");
            if (type == null)
                throw new ArgumentNullException("type");

            if (memberInfo.DeclaringType == type)
                return 0;
            if (type.BaseType == null)
                throw new ArgumentException("memberInfo is not declared on type or anywhere in its inheritance tree");
            return GetMemberInheritanceDepth(memberInfo, type.BaseType) + 1;
        }

        private Predicate<string> GetNameMatcher(string name, MemberNameMatchBehaviourOptions memberNameMatchBehaviour)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            if (!Enum.IsDefined(typeof(MemberNameMatchBehaviourOptions), memberNameMatchBehaviour))
                throw new ArgumentOutOfRangeException("memberNameMatchBehaviour");

            if (memberNameMatchBehaviour == MemberNameMatchBehaviourOptions.CaseInsensitive)
                return n => n.Equals(name, StringComparison.InvariantCultureIgnoreCase);
            if (memberNameMatchBehaviour == MemberNameMatchBehaviourOptions.Precise)
                return n => n.Equals(name, StringComparison.InvariantCultureIgnoreCase);
            if (memberNameMatchBehaviour == MemberNameMatchBehaviourOptions.UseNameRewriter)
            {
                var rewritterName = _nameRewriter(name);
                return n => n == rewritterName;
            }
            throw new ArgumentOutOfRangeException("memberNameMatchBehaviour");
        }

        private enum DefaultMemberBehaviourOptions
        {
            DoesNotMatter,
            MustBeDefault
        }

        private enum MemberNameMatchBehaviourOptions
        {
            CaseInsensitive,
            Precise,
            UseNameRewriter
        }

        private bool TypeIsTranslatedFromVBScript(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            return type.GetCustomAttributes(typeof(SourceClassName), true).Any();
        }

        private bool TypeIsComVisible(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            // For a type to be considered ComVisible, there must be a ComVisible(true) attribute on it or one of its explicitly derived types, having it
            // only on object (which everything implicitly derives from) is not sufficient since then everything would be ComVisible (since object has
            // the ComVisible(true) attribute set)
            if (type == typeof(object))
                return false;

            // We need to consider the ComVisible attributes on every class in the hierarchy (other than object - see above), so we specify inherit: false
            // here and walk up the tree until we find a ComVisible(true) - any ComVisible(false) settings encountered can be ignored so long as one of
            // the classes in the hierarchy has ComVisible(true).
            if (type.GetCustomAttributes(typeof(ComVisibleAttribute), inherit: false).Cast<ComVisibleAttribute>().Any(a => a.Value))
                return true;

            if (type.BaseType == null)
                return false;

            return TypeIsComVisible(type.BaseType);
        }

        private bool MemberHasDispIdZero(MemberInfo memberInfo)
        {
            if (memberInfo == null)
                throw new ArgumentNullException("memberInfo");

            return memberInfo.GetCustomAttributes(typeof(DispIdAttribute), inherit: true)
                .Cast<DispIdAttribute>()
                .Any(attribute => attribute.Value == 0);
        }

        private bool AnyDispIdZeroMemberExists(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            return type.GetMembers().Any(m => MemberHasDispIdZero(m));
        }

        private bool DispIdZeroIsAmbiguous(Type type)
        {
            if (type == null)
                throw new ArgumentNullException("type");

            return type.GetMembers().Count(m => MemberHasDispIdZero(m)) > 1;
        }

        /// <summary>
        /// VBScript considers all of these to be value types (nulls and string as well as CLR value types)
        /// </summary>
        private bool IsVBScriptValueType(object o)
        {
            return ((o == null) || (o == DBNull.Value) || (o is ValueType) || (o is string));
        }

        private sealed class InvokerCacheKey
        {
            private readonly int _hashCode;
            public InvokerCacheKey(object targetType, string optionalName, int numberOfArguments)
            {
                if (targetType == null)
                    throw new ArgumentNullException("targetType");
                if (numberOfArguments < 0)
                    throw new ArgumentOutOfRangeException("numberOfArguments", "must be zero or greater");

                TargetType = targetType;
                OptionalName = optionalName;
                NumberOfArguments = numberOfArguments;

                _hashCode = (TargetType.ToString() + "\n" + (optionalName ?? "") + "\n" + numberOfArguments.ToString()).GetHashCode();
            }

            /// <summary>
            /// This will never be null
            /// </summary>
            public object TargetType { get; private set; }

            /// <summary>
            /// This is optional and may be null
            /// </summary>
            public string OptionalName { get; private set; }

            /// <summary>
            /// This will always be zero or greater
            /// </summary>
            public int NumberOfArguments { get; private set; }

            public override int GetHashCode()
            {
                return _hashCode;
            }

            public override bool Equals(object obj)
            {
                if (obj == null)
                    throw new ArgumentNullException("obj");
                var cacheKey = obj as InvokerCacheKey;
                if (cacheKey == null)
                    return false;
                return (
                    (TargetType == cacheKey.TargetType) &&
                    (OptionalName == cacheKey.OptionalName) &&
                    (NumberOfArguments == cacheKey.NumberOfArguments)
                );
            }
        }

        private class ManagedEnumeratorWrapper : IEnumerable
        {
            private readonly IEnumerator _enumerator;
            public ManagedEnumeratorWrapper(IEnumerator enumerator)
            {
                if (enumerator == null)
                    throw new ArgumentNullException("enumerator");

                _enumerator = enumerator;
            }

            public IEnumerator GetEnumerator()
            {
                return _enumerator;
            }
        }

        private class IDispatchEnumeratorWrapper : IEnumerator
        {
            private readonly IEnumVariant _enumerator;
            private object _current;
            public IDispatchEnumeratorWrapper(IEnumVariant enumerator)
            {
                if (enumerator == null)
                    throw new ArgumentNullException("enumerator");

                _enumerator = enumerator;
                _current = null;
            }

            public object Current { get { return _current; } }

            public bool MoveNext()
            {
                uint fetched;
                _enumerator.Next(1, out _current, out fetched);
                return fetched != 0;
            }

            public void Reset()
            {
                _enumerator.Reset();
            }
        }

        [ComImport(), Guid("00020404-0000-0000-C000-000000000046"), InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumVariant
        {
            // From https://github.com/mosa/Mono-Class-Libraries/blob/master/mcs/class/CustomMarshalers/System.Runtime.InteropServices.CustomMarshalers/EnumeratorToEnumVariantMarshaler.cs
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Next(int celt, [MarshalAs(UnmanagedType.Struct)]out object rgvar, out uint pceltFetched);
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Skip(uint celt);
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Reset();
            [return: MarshalAs(UnmanagedType.Interface)]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            IEnumVariant Clone();
        }
    }
}
