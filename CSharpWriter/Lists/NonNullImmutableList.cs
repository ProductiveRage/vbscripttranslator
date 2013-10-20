using System;
using System.Collections.Generic;

namespace CSharpWriter.Lists
{
    public class NonNullImmutableList<T> : ImmutableList<T> where T : class
    {
        private readonly static Validator _defaultValidator = new Validator(null);
        private IValueValidator<T> _optionalValueValidator;

        public NonNullImmutableList() : this((IValueValidator<T>)null) { }
        public NonNullImmutableList(IEnumerable<T> values) : this(values, null) { }
        public NonNullImmutableList(IValueValidator<T> optionalValueValidator)
            : base((Node)null, GetValidator(optionalValueValidator))
        {
            _optionalValueValidator = optionalValueValidator;
        }
        public NonNullImmutableList(IEnumerable<T> values, IValueValidator<T> optionalValueValidator)
            : base(values, GetValidator(optionalValueValidator))
        {
            _optionalValueValidator = optionalValueValidator;
        }
        private NonNullImmutableList(Node tail, IValueValidator<T> optionalValueValidator)
            : base(tail, GetValidator(optionalValueValidator))
        {
            _optionalValueValidator = optionalValueValidator;
        }

        private static IValueValidator<T> GetValidator(IValueValidator<T> optionalValueValidator)
        {
            if (optionalValueValidator == null)
                return _defaultValidator;
            return new Validator(optionalValueValidator);
        }

        public new NonNullImmutableList<T> Add(T value)
        {
            return ToNonNullImmutableList(base.Add(value));
        }
        public new NonNullImmutableList<T> AddRange(IEnumerable<T> values)
        {
            return ToNonNullImmutableList(base.AddRange(values));
        }
        public new NonNullImmutableList<T> Insert(T value, int insertAtIndex)
        {
            return ToNonNullImmutableList(base.Insert(value, insertAtIndex));
        }
        public new NonNullImmutableList<T> Insert(IEnumerable<T> values, int insertAtIndex)
        {
            return ToNonNullImmutableList(base.Insert(values, insertAtIndex));
        }
        public new NonNullImmutableList<T> Remove(T value)
        {
            return ToNonNullImmutableList(base.Remove(value));
        }
        public new NonNullImmutableList<T> Remove(T value, IEqualityComparer<T> optionalComparer)
        {
            return ToNonNullImmutableList(base.Remove(value, optionalComparer));
        }
        public new NonNullImmutableList<T> RemoveAt(int removeAtIndex)
        {
            return ToNonNullImmutableList(base.RemoveAt(removeAtIndex));
        }
        public new NonNullImmutableList<T> RemoveRange(int removeAtIndex, int count)
        {
            return ToNonNullImmutableList(base.RemoveRange(removeAtIndex, count));
        }
        public new NonNullImmutableList<T> RemoveLast()
        {
            return ToNonNullImmutableList(base.RemoveLast());
        }
        public new NonNullImmutableList<T> RemoveLast(int numberToRemove)
        {
            return ToNonNullImmutableList(base.RemoveLast(numberToRemove));
        }
        public new NonNullImmutableList<T> Sort()
        {
            return ToNonNullImmutableList(base.Sort());
        }
        public new NonNullImmutableList<T> Sort(Comparison<T> optionalComparison)
        {
            return ToNonNullImmutableList(base.Sort(optionalComparison));
        }
        public new NonNullImmutableList<T> Sort(IComparer<T> optionalComparer)
        {
            return ToNonNullImmutableList(base.Sort(optionalComparer));
        }
        private NonNullImmutableList<T> ToNonNullImmutableList(ImmutableList<T> list)
        {
            if (list == null)
                throw new ArgumentNullException("list");

            return To<NonNullImmutableList<T>>(
                list,
                tail => new NonNullImmutableList<T>(tail, _optionalValueValidator)
            );
        }

        private class Validator : IValueValidator<T>
        {
            private IValueValidator<T> _optionalInnerValidator;
            public Validator(IValueValidator<T> optionalInnerValidator)
            {
                _optionalInnerValidator = optionalInnerValidator;
            }

            /// <summary>
            /// This will throw an exception for a value that does pass validation requirements
            /// </summary>
            public void EnsureValid(T value)
            {
                if (value == null)
                    throw new ArgumentNullException("value");
                if (_optionalInnerValidator != null)
                    _optionalInnerValidator.EnsureValid(value);
            }
        }
    }
}
