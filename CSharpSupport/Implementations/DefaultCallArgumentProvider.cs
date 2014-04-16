using System;
using System.Collections.Generic;
using System.Linq;

namespace CSharpSupport.Implementations
{
    public class DefaultCallArgumentProvider : IBuildCallArgumentProviders
    {
        private readonly List<Tuple<object, Action<object>>> _valuesWithUpdatesWhereRequired;
		private readonly IAccessValuesUsingVBScriptRules _vbscriptValueAccessor;
        public DefaultCallArgumentProvider(IAccessValuesUsingVBScriptRules vbscriptValueAccessor)
        {
			if (vbscriptValueAccessor == null)
				throw new ArgumentNullException("vbscriptValueAccessor");

			_vbscriptValueAccessor = vbscriptValueAccessor;
			_valuesWithUpdatesWhereRequired = new List<Tuple<object, Action<object>>>();
        }

        /// <summary>
        /// TODO
        /// This should return a reference to itself to enable chaining when building up argument sets
        /// </summary>
        public IBuildCallArgumentProviders Val(object value)
        {
            _valuesWithUpdatesWhereRequired.Add(Tuple.Create(value, (Action<object>)null));
            return this;
        }

        /// <summary>
        /// TODO
        /// This should return a reference to itself to enable chaining when building up argument sets
        /// </summary>
        public IBuildCallArgumentProviders Ref(object value, Action<object> valueUpdater)
        {
            if (valueUpdater == null)
                throw new ArgumentNullException("valueUpdater");

            _valuesWithUpdatesWhereRequired.Add(Tuple.Create(value, valueUpdater));
            return this;
        }

        /// <summary>
        /// TODO
        /// This should return a reference to itself to enable chaining when building up argument sets
        /// </summary>
        public IBuildCallArgumentProviders RefIfArray(object target, IEnumerable<IProvideCallArguments> argumentProviders)
        {
            if (target == null)
                throw new ArgumentNullException("target");
			if (argumentProviders == null)
				throw new ArgumentNullException("argumentProviders");

			var argumentProvidersArray = argumentProviders.ToArray();
			if (argumentProvidersArray.Length == 0)
				throw new ArgumentException("There must be at least one argument provider");
			if (argumentProvidersArray.Any(p => p.NumberOfArguments == 0))
				throw new ArgumentException("There may not be any argument providers with zero arguments");

			// Process all but the last set of argument providers, updating target with each call. If at any point target is not an array
			// then the final value will be passed ByVal (since there must be a function or property access involved, the result of which
			// is never passed ByRef).
			var passByVal = false;
			for (var index = 0; index < argumentProvidersArray.Length - 1; index++)
			{
				if (!target.GetType().IsArray)
					passByVal = true;
				target = _vbscriptValueAccessor.CALL(target, argumentProvidersArray[index]);
			}
			if (!target.GetType().IsArray)
				passByVal = true;

			// Process the final arguments to get the value that should actually be passed as the argument. If we've determined that this
			// value should be passed ByVal then hand straight off to the Val method.
			var lastArgumentProvider = argumentProvidersArray.Last();
			var valueForArgument = _vbscriptValueAccessor.CALL(target, lastArgumentProvider);
			if (passByVal)
				return Val(valueForArgument);

			// If the value must be passed ByRef then we pass in the same valueForArgument as above but in the valueUpdater callback we
			// have to call SET on the target (which is the array that valueForArgument was taken from) to push the ByRef value back
			// into the array. (The ByRef support here is only used if the method being called has ByRef arguments, otherwise it's
			// not required - but that is the responsibility of the CALL implementation to write the values back, the job of this
			// interface is to allow that to occur where necessary).
			return Ref(
				valueForArgument,
				v => _vbscriptValueAccessor.SET(target, null, lastArgumentProvider.GetInitialValues().ToArray(), v)
			);
        }

        /// <summary>
        /// TODO
        /// </summary>
        public IProvideCallArguments GetArgs()
        {
            return new ArgumentProvider(_valuesWithUpdatesWhereRequired);
        }

        private class ArgumentProvider : IProvideCallArguments
        {
            private readonly List<Tuple<object, Action<object>>> _valuesWithUpdatesWhereRequired;
            public ArgumentProvider(List<Tuple<object, Action<object>>> valuesWithUpdatesWhereRequired)
            {
                if (valuesWithUpdatesWhereRequired == null)
                    throw new ArgumentNullException("valuesWithUpdatesWhereRequired");

                _valuesWithUpdatesWhereRequired = valuesWithUpdatesWhereRequired;
            }

            /// <summary>
            /// This will always be zero or greater
            /// </summary>
            public int NumberOfArguments { get { return _valuesWithUpdatesWhereRequired.Count; } }

            /// <summary>
            /// This will always return a set with NumberOfArguments items in it
            /// </summary>
            public IEnumerable<object> GetInitialValues()
            {
                return _valuesWithUpdatesWhereRequired.Select(entry => entry.Item1);
            }

            /// <summary>
            /// The index must be zero or greater and less than NumberOfArguments. If the argument at that index may not be overrwritten then the
            /// function call will have no effect.
            /// </summary>
            public void OverwriteValueIfByRef(int index, object value)
            {
                if ((index < 0) || (index >= _valuesWithUpdatesWhereRequired.Count))
                    throw new ArgumentOutOfRangeException("index");

                var valueUpdater = _valuesWithUpdatesWhereRequired[index].Item2;
                if (valueUpdater != null)
                    valueUpdater(value);
            }
        }
    }
}
