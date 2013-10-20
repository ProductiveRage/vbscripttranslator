using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;

namespace CSharpWriter.Lists
{
    public class ImmutableList<T> : IEnumerable<T>
    {
        private readonly Node _tail;
        private readonly IValueValidator<T> _optionalValueValidator;
        private T[] _allValues;

        public ImmutableList() : this((IValueValidator<T>)null) { }
        public ImmutableList(IEnumerable<T> values) : this(values, null) { }
        public ImmutableList(IValueValidator<T> optionalValueValidator) : this((Node)null, optionalValueValidator) { }
        public ImmutableList(IEnumerable<T> values, IValueValidator<T> optionalValueValidator)
        {
            if (values == null)
                throw new ArgumentNullException("values");

            Node node = null;
            foreach (var value in values)
            {
                if (optionalValueValidator != null)
                    optionalValueValidator.EnsureValid(value);
                if (node == null)
                    node = new Node(value, null);
                else
                    node = new Node(value, node);
            }
            _tail = node;
            _optionalValueValidator = optionalValueValidator;
            _allValues = null;
        }
        protected ImmutableList(Node tail, IValueValidator<T> optionalValueValidator)
        {
            _tail = tail;
            _optionalValueValidator = optionalValueValidator;
            _allValues = null;
        }

        public T this[int index]
        {
            get
            {
                if ((index < 0) || (index >= Count))
                    throw new ArgumentOutOfRangeException("index");

                // Getting the value of the last item is a very quick operation so we can add a shortcut for it
                if (index == Count - 1)
                    return _tail.Value;

                EnsureAllValuesDataIsPopulated();
                return _allValues[index];
            }
        }

        public int Count
        {
            get { return (_tail == null) ? 0 : _tail.Count; }
        }

        public bool Contains(T value)
        {
            return Contains(value, null);
        }

        public bool Contains(T value, IEqualityComparer<T> optionalComparer)
        {
            if (_tail == null)
                return false;

            EnsureAllValuesDataIsPopulated();
            for (var index = 0; index < _allValues.Length; index++)
            {
                if (DoValuesMatch(_allValues[index], value, optionalComparer))
                    return true;
            }
            return false;
        }

        public ImmutableList<T> Add(T value)
        {
            // Add is easy since we keep a reference to the tail node, we only need to wrap it in a new node to create a new tail!
            if (_optionalValueValidator != null)
                _optionalValueValidator.EnsureValid(value);
            return new ImmutableList<T>(
                new Node(value, _tail),
                _optionalValueValidator
            );
        }

        public ImmutableList<T> AddRange(IEnumerable<T> values)
        {
            if (values == null)
                throw new ArgumentNullException("values");
            if (!values.Any())
                return this;

            // AddRange is easy for the same reason as Add
            var node = _tail;
            foreach (var value in values)
            {
                if (_optionalValueValidator != null)
                    _optionalValueValidator.EnsureValid(value);
                node = new Node(value, node);
            }
            return new ImmutableList<T>(node, _optionalValueValidator);
        }

        public ImmutableList<T> Insert(IEnumerable<T> values, int insertAtIndex)
        {
            if (values == null)
                throw new ArgumentNullException("values");

            return Insert(values, default(T), insertAtIndex);
        }

        public ImmutableList<T> Insert(T value, int insertAtIndex)
        {
            return Insert(null, value, insertAtIndex);
        }

        private ImmutableList<T> Insert(IEnumerable<T> multipleValuesToAdd, T singleValueToAdd, int insertAtIndex)
        {
            if ((insertAtIndex < 0) || (insertAtIndex > Count))
                throw new ArgumentOutOfRangeException("insertAtIndex");
            if ((multipleValuesToAdd != null) && !multipleValuesToAdd.Any())
                return this;

            // If the insertion is at the end of the list then we can use Add or AddRange which may allow some optimisation
            if (insertAtIndex == Count)
            {
                if (multipleValuesToAdd == null)
                    return Add(singleValueToAdd);
                return AddRange(multipleValuesToAdd);
            }

            // Starting with the tail, walk back to the insertion point, record the values we pass over
            var node = _tail;
            var valuesBeforeInsertionPoint = new T[Count - insertAtIndex];
            for (var index = 0; index < valuesBeforeInsertionPoint.Length; index++)
            {
                valuesBeforeInsertionPoint[index] = node.Value;
                node = node.Previous;
            }

            // Any existing node chain before the insertion point can be persisted and the new value(s) appended
            if (multipleValuesToAdd == null)
            {
                if (_optionalValueValidator != null)
                    _optionalValueValidator.EnsureValid(singleValueToAdd);
                node = new Node(singleValueToAdd, node);
            }
            else
            {
                foreach (var valueToAdd in multipleValuesToAdd)
                {
                    if (_optionalValueValidator != null)
                        _optionalValueValidator.EnsureValid(valueToAdd);
                    node = new Node(valueToAdd, node);
                }
            }

            // Finally, toAdd back the values we walked through before to complete the chain
            for (var index = valuesBeforeInsertionPoint.Length - 1; index >= 0; index--)
                node = new Node(valuesBeforeInsertionPoint[index], node);
            return new ImmutableList<T>(node, _optionalValueValidator);
        }

        /// <summary>
        /// Removes the first occurrence of a specific object from the list, if the item is not present then this instance will be returned
        /// </summary>
        public ImmutableList<T> Remove(T value)
        {
            return Remove(value, null);
        }

        /// <summary>
        /// Removes the first occurrence of a specific object from the list, if the item is not present then this instance will be returned
        /// </summary>
        public ImmutableList<T> Remove(T value, IEqualityComparer<T> optionalComparer)
        {
            // If there are no items in the list then the specified value can't be present, so do nothing
            if (_tail == null)
                return this;

            // Try to find the last node that matches the value when walking backwards from the tail; this will be the first in the list
            // when considered from start to end
            var node = _tail;
            Node lastNodeThatMatched = null;
            int? lastNodeIndexThatMatched = null;
            var valuesBeforeRemoval = new T[Count];
            for (var index = 0; index < Count; index++)
            {
                if (DoValuesMatch(value, node.Value, optionalComparer))
                {
                    lastNodeThatMatched = node;
                    lastNodeIndexThatMatched = index;
                }
                valuesBeforeRemoval[index] = node.Value;
                node = node.Previous;
            }
            if (lastNodeThatMatched == null)
                return this;

            // Now build a new chain by taking the content before the value-to-remove and adding back the values that were stepped through
            node = lastNodeThatMatched.Previous;
            for (var index = lastNodeIndexThatMatched.Value - 1; index >= 0; index--)
                node = new Node(valuesBeforeRemoval[index], node);
            return new ImmutableList<T>(node, _optionalValueValidator);
        }

        public ImmutableList<T> Sort()
        {
            return Sort((IComparer<T>)null);
        }

        public ImmutableList<T> Sort(Comparison<T> optionalComparison)
        {
            if (optionalComparison == null)
                return Sort((IComparer<T>)null);
            return Sort(new SortComparisonWrapper(optionalComparison));
        }

        public ImmutableList<T> Sort(IComparer<T> optionalComparer)
        {
            EnsureAllValuesDataIsPopulated();
            return new ImmutableList<T>(
                (optionalComparer == null) ? _allValues.OrderBy(v => v) : _allValues.OrderBy(v => v, optionalComparer),
                _optionalValueValidator
            );
        }

        private bool DoValuesMatch(T x, T y, IEqualityComparer<T> optionalComparer)
        {
            if (optionalComparer != null)
                return optionalComparer.Equals(x, y);

            if ((x == null) && (y == null))
                return true;
            else if ((x == null) || (y == null))
                return false;
            else
                return x.Equals(y);
        }

        public ImmutableList<T> RemoveAt(int removeAtIndex)
        {
            return RemoveRange(removeAtIndex, 1);
        }

        public ImmutableList<T> RemoveRange(int removeAtIndex, int count)
        {
            if (removeAtIndex < 0)
                throw new ArgumentOutOfRangeException("removeAtIndex", "must be greater than or equal zero");
            if (count <= 0)
                throw new ArgumentOutOfRangeException("count", "must be greater than zero");
            if ((removeAtIndex + count) > Count)
                throw new ArgumentException("removeAtIndex + count must not exceed Count");

            // Starting with the tail, walk back to the end of the removal range, recording the values we pass over
            var node = _tail;
            var valuesBeforeRemovalRange = new T[Count - (removeAtIndex + count)];
            for (var index = 0; index < valuesBeforeRemovalRange.Length; index++)
            {
                valuesBeforeRemovalRange[index] = node.Value;
                node = node.Previous;
            }

            // Move past the values in the removal range
            for (var index = 0; index < count; index++)
                node = node.Previous;

            // Now toAdd back the values we walked through above to the part of the chain that can be persisted
            for (var index = valuesBeforeRemovalRange.Length - 1; index >= 0; index--)
                node = new Node(valuesBeforeRemovalRange[index], node);
            return new ImmutableList<T>(node, _optionalValueValidator);
        }

        public ImmutableList<T> RemoveLast()
        {
            return RemoveLast(1);
        }

        public ImmutableList<T> RemoveLast(int numberToRemove)
        {
            if (numberToRemove <= 0)
                throw new ArgumentOutOfRangeException("numberToRemove", "must be greater than zero");
            if (numberToRemove > Count)
                throw new ArgumentOutOfRangeException("numberToRemove", "must not be greater than Count");

            var node = _tail;
            for (var index = 0; index < numberToRemove; index++)
                node = node.Previous;
            return new ImmutableList<T>(node, _optionalValueValidator);
        }

        public IEnumerator<T> GetEnumerator()
        {
            // As documented at http://msdn.microsoft.com/en-us/library/system.array.aspx, from .Net 2.0 onward, the Array class implements
            // IEnumerable<T> but this is only provided at runtime so we have to explicitly cast access its generic GetEnumerator method
            EnsureAllValuesDataIsPopulated();
            return ((IEnumerable<T>)_allValues).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// So that a derived class may override the public methods with implementations that return the derived type's class, this method exposes a manner
        /// to access the _tail reference of a return ImmutableList instance without having to make both it and the Node class public - eg. a derived class
        /// NonNullOrEmptyStringList may incorporate its own hard-coded validation and wish to have a NonNullOrEmptyStringList instance returned from its
        /// Add method. If it calls the ImmutableList's Add method it will receive a new ImmutableList instance which can be transformed into an instance
        /// of NonNullOrEmptyStringList if it has a constructor which take a Node argument by passing a lambda wrapping a call to that constructor into
        /// this method, along with the new ImmutableList reference that is to be wrapped. This introduce does have the overhead of an additional
        /// initialisation (of the NonNullOrEmptyStringList) but it allows for more strictly-typed return values from the NonNullOrEmptyStringList's
        /// methods.
        /// </summary>
        protected static U To<U>(ImmutableList<T> list, Func<Node, U> generator)
        {
            if (list == null)
                throw new ArgumentNullException("list");
            if (generator == null)
                throw new ArgumentNullException("generator");

            return generator(list._tail);
        }

        /// <summary>
        /// For enumerating the values we need to walk through all of the nodes and then reverse the set (since we start with the tail and work backwards).
        /// This can be relatively expensive so the list is cached in the "_allValues" member array so that subsequent requests are fast (wouldn't be a
        /// big deal for a single enumeration of the contents but it could be for multiple calls to the indexed property).
        /// </summary>
        private void EnsureAllValuesDataIsPopulated()
        {
            if (_allValues != null)
                return;

            // Since we start at the tail and work backwards, we need to reverse the order of the items in values array that is populated here
            var numberOfValues = Count;
            var values = new T[numberOfValues];
            var node = _tail;
            for (var index = 0; index < numberOfValues; index++)
            {
                values[(numberOfValues - 1) - index] = node.Value;
                node = node.Previous;
            }
            _allValues = values;
        }

        /// <summary>
        /// This is used by the Sort method if a Comparison<T> is specified
        /// </summary>
        private class SortComparisonWrapper : IComparer<T>
        {
            private Comparison<T> _comparison;
            public SortComparisonWrapper(Comparison<T> comparison)
            {
                if (comparison == null)
                    throw new ArgumentNullException("comparison");

                _comparison = comparison;
            }

            public int Compare(T x, T y)
            {
                return _comparison(x, y);
            }
        }

        protected class Node
        {
            public Node(T value, Node previous)
            {
                Value = value;
                Previous = previous;
                Count = (previous == null) ? 1 : (previous.Count + 1);
            }

            public T Value { get; private set; }

            /// <summary>
            /// This will be null if there is no previous node (ie. this is the start of the chain, the head)
            /// </summary>
            public Node Previous { get; private set; }

            public int Count { get; private set; }
        }
    }
}
