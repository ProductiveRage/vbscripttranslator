﻿using CSharpSupport.Attributes;
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
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
		private readonly ConcurrentDictionary<InvokerCacheKey, GetInvoker> _invokerCache;
		public VBScriptEsqueValueRetriever(Func<string, string> nameRewriter)
        {
            if (nameRewriter == null)
                throw new ArgumentNullException("nameRewriter");

            _nameRewriter = nameRewriter;
			_invokerCache = new ConcurrentDictionary<InvokerCacheKey, GetInvoker>();
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
        /// This requires a target with optional member accessors and arguments - eg. "Test" is a target only, "a.Test" has target "a" with one
        /// named member "Test", "a.Test(0)" has target "a", named member "Test" and a single argument "0". The expression "a(Test(0))" would
        /// require nested CALL executions, one with target "Test" and a single argument "0" and a second with target "a" and a single
        /// argument which was the result of the first call. A null target is only acceptable if there are no member or arguments
        /// specified, in which null will be returned.
        /// </summary>
        public object CALL(object target, IEnumerable<string> members, params object[] arguments)
        {
            if (members == null)
                throw new ArgumentNullException("members");
            if (arguments == null)
                throw new ArgumentNullException("arguments");

            var memberAccessorsArray = members.ToArray();
            if (memberAccessorsArray.Any(m => string.IsNullOrWhiteSpace(m)))
                throw new ArgumentException("Null/blank value in members set");

            var argumentsArray = arguments.ToArray();
            if (argumentsArray.Any(a => a == null))
                throw new ArgumentException("Null reference encountered in arguments set");

            // If there are no member accessor or arguments then just return the object directly (if the caller wants this to be a value type
            // then they'll have to use VAL to try to apply that requirement). This is the only case in which it's acceptable for target to
            // be null.
            if (!memberAccessorsArray.Any() && !argumentsArray.Any())
                return target;
            else if (target == null)
                throw new ArgumentException("The target reference may only be null if there are no member accessors or arguments specified");

            // If there are no member accessors but there are arguments then
            // 1. Attempt array access (this is the only time that array access is acceptable (o.Names(0) is not allowed by VBScript if the
            //    "Names" property is an array, it will only work if Names if a Function or indexed Property)
            // 2. If target implements IDispatch then try accessing the DispId zero method or property, passing the arguments
            // 3. If it's not an IDispatch reference, then the arguments will be passed to a method or indexed property that will accept them,
            //    taking into account the IsDefault attribute
            if (!memberAccessorsArray.Any() && argumentsArray.Any())
                return InvokeGetter(target, null, argumentsArray);

            // If there are member accessors but no arguments then we walk down each member accessor, no defaults are considered
            if (!argumentsArray.Any())
                return WalkMemberAccessors(target, memberAccessorsArray);

            // If there member accessors AND arguments then all-but-the-last member accessors should be walked through as argument-less lookups
            // and the final member accessor should be a method or property call whose name matches the member accessor. Note that the arguments
            // can never be for an array look up if there are member accessors, it must always be a function or property.
            target = WalkMemberAccessors(target, memberAccessorsArray.Take(memberAccessorsArray.Length - 1));
            var finalMemberAccessor = memberAccessorsArray[memberAccessorsArray.Length - 1];
            if (target == null)
                throw new ArgumentException("Unable to access member \"" + finalMemberAccessor + "\" on null reference");
            return InvokeGetter(target, finalMemberAccessor, argumentsArray);
        }


		private object InvokeGetter(object target, string optionalName, IEnumerable<object> arguments)
		{
			if (target == null)
				throw new ArgumentNullException("target");
			if (arguments == null)
				throw new ArgumentNullException("arguments");

			var argumentsArray = arguments.ToArray();
			var cacheKey = new InvokerCacheKey(target.GetType(), optionalName, argumentsArray.Length);
			GetInvoker invoker;
			if (!_invokerCache.TryGetValue(cacheKey, out invoker))
			{
				invoker = GenerateGetInvoker(target, optionalName, argumentsArray);
				_invokerCache.TryAdd(cacheKey, invoker);
			}
			return invoker(target, argumentsArray);
		}

		private delegate object GetInvoker(object target, object[] arguments);
		private GetInvoker GenerateGetInvoker(object target, string optionalName, IEnumerable<object> arguments)
		{
			if (target == null)
				throw new ArgumentNullException("target");

            var argumentsArray = arguments.ToArray();
            var targetType = target.GetType();
            if (targetType.IsArray)
            {
                if (targetType.GetArrayRank() != argumentsArray.Length)
                    throw new ArgumentException("Argument count (" + argumentsArray.Length + ") does not match arrary rank (" + targetType.GetArrayRank() + ")");

                // TODO: Incorporate crazy VBScript logic into argument casting
                var arrayTargetParameter = Expression.Parameter(typeof(object), "target");
                var indexesParameter = Expression.Parameter(typeof(object[]), "arguments");
                var arrayAccessExceptionParameter = Expression.Parameter(typeof(Exception), "e");
                return Expression.Lambda<GetInvoker>(
				    Expression.TryCatch(
                        Expression.Convert(
                            Expression.ArrayAccess(
                                Expression.Convert(arrayTargetParameter, targetType),
                                Enumerable.Range(0, argumentsArray.Length).Select(index =>
							        Expression.Convert(
								        Expression.ArrayAccess(indexesParameter, Expression.Constant(index)),
								        typeof(int)
							        )
						        )
                            ),
                            typeof(object) // Without this we may get an "Expression of type 'System.Int32' cannot be used for return type 'System.Object'" or similar
                        ),
                        Expression.Catch(
						    arrayAccessExceptionParameter,
                            Expression.Throw(
                                GetNewArgumentException("Error accessing array with specified indexes (likely a non-numeric array index/argument)", arrayAccessExceptionParameter),
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
			return Expression.Lambda<GetInvoker>(

				Expression.TryCatch(

					Expression.Call(
                        Expression.Convert(targetParameter, targetType),
						method,
						method.GetParameters().Select((arg, index) =>
							Expression.Convert(
								Expression.ArrayAccess(argumentsParameter, Expression.Constant(index)),
								arg.ParameterType
							)
						)
					),

                    // The Throw requires a return type to be specified since the Try block has a return type - without
                    // this a runtime "Body of catch must have the same type as body of try" exception will be raised
                    Expression.Catch(
						exceptionParameter,
                        Expression.Throw(
                            GetNewArgumentException("Error executing " + errorMessageMemberDescription, exceptionParameter),
							typeof(object)
						)
					)

				),
				new[]
				{
					targetParameter,
					argumentsParameter
				}
			).Compile();
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
            // TODO: Should all ComVisible classes that this might apply to implement IDispatch and so make this unnecessary?
            return GetGetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.UseNameRewriter, numberOfArguments)
                .Concat(
                    GetGetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.Precise, numberOfArguments)
                )
                .Concat(
                    GetGetMethods(type, name, DefaultMemberBehaviourOptions.DoesNotMatter, MemberNameMatchBehaviourOptions.CaseInsensitive, numberOfArguments)
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
