using System;
using System.Collections.Generic;
using System.Linq;

namespace CSharpSupport.Implementations
{
    public class DefaultCallArgumentProvider : IBuildCallArgumentProviders
    {
        private readonly List<Tuple<object, Action<object>>> _valuesWithUpdatesWhereRequired;
        public DefaultCallArgumentProvider()
        {
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
        public IBuildCallArgumentProviders RefIfArray(object target, IEnumerable<IBuildCallArgumentProviders> argumentProviders)
        {
            if (target == null)
                throw new ArgumentNullException("target");
			if (argumentProviders == null)
				throw new ArgumentNullException("argumentProviders");

            throw new NotImplementedException(); // TODO
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
