using CSharpSupport.Attributes;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

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

            if (o == null)
                throw new Exception("Object expected (failed trying to extract value type data from null reference)");

            var defaultValueFromObject = InvokeGetter(o, null, new object[0]);
            if (IsVBScriptValueType(defaultValueFromObject))
                return defaultValueFromObject;

            // We don't recursively try defaults, so if this default is still not a value type then we're out of luck
            throw new Exception("Object expected (default method/property of object also returned non-value type data)");
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
        /// Reduce a reference down to a boolean, throwing an exception if this is not possible. This will apply the same logic as VAL but then
        /// require a numeric value or null, otherwise an exception will be raised. Zero and null equate to false, non-zero numbers to true.
        /// </summary>
        public bool IF(object o)
        {
            if (o == null)
                return false;

            var valueString = VAL(o).ToString();

            double value;
            if (!double.TryParse(valueString, out value))
                throw new ArgumentException("Type Mismatch: [string \"" + valueString + "\"] (unable to translate into boolean for IF statement)");

            return value != 0;
        }

        /// <summary>
        /// TODO
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
        /// possible (in that case, a straight assignment should be generated - no call to SET required).
        /// </summary>
        public void SET(object target, string optionalMemberAccessor, IProvideCallArguments argumentProvider, object value)
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
            invoker(target, arguments, value);
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
                        throw new ArgumentException("Unable to identify " + errorMessageMemberDescription + " (target implements IDispatch)", e);
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
                throw new ArgumentException("Unable to identify " + errorMessageMemberDescription);

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

            // Note: Considered filtering out methods that have a void ReturnType but then we wouldn't be able to call functions that originated from
            // VBScript SUBSs, which would be a problem! Null will be returned if the method is successfully invoked, if the method's ReturnType is
            // void. Valid source script would never try to access the value returned since a SUB never returns a value.
            var nameMatcher = (optionalName != null) ? GetNameMatcher(optionalName, memberNameMatchBehaviour) : (name => true);
            return
                type.GetMethods()
                    .Where(m => nameMatcher(m.Name))
                    .Where(m => m.GetParameters().Length == numberOfArguments)
                    .Where(m =>
                        (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                        (m.GetCustomAttributes(true).Cast<Attribute>().Any(a => a is IsDefault))
                    )
                .Concat(
                    type.GetProperties()
                        .Where(p => p.CanRead)
                        .Where(p => nameMatcher(p.Name))
                        .Where(p => p.GetIndexParameters().Length == numberOfArguments)
                        .Where(p =>
                            (defaultMemberBehaviour == DefaultMemberBehaviourOptions.DoesNotMatter) ||
                            (p.GetCustomAttributes(true).Cast<Attribute>().Any(a => a is IsDefault))
                        )
                        .Select(p => p.GetGetMethod())
                );
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
                type.GetMethods()
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

        /// <summary>
        /// VBScript considers all of these to be value types (nulls and string as well as CLR value types)
        /// </summary>
        private bool IsVBScriptValueType(object o)
        {
            return ((o == null) || (o is ValueType) || (o is string));
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
    }
}
