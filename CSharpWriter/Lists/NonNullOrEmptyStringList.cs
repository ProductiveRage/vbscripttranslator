using System;
using System.Collections.Generic;

namespace CSharpWriter.Lists
{
    public class NonNullOrEmptyStringList : ImmutableList<string>
    {
        private readonly static Validator _validator = new Validator();

        public NonNullOrEmptyStringList() : this((Node)null) { }
        public NonNullOrEmptyStringList(IEnumerable<string> values) : base(values, _validator) { }
        private NonNullOrEmptyStringList(Node tail) : base(tail, _validator) { }

        public new NonNullOrEmptyStringList Add(string value)
        {
            return ToNonNullOrEmptyStringList(base.Add(value));
        }
        public new NonNullOrEmptyStringList AddRange(IEnumerable<string> values)
        {
            return ToNonNullOrEmptyStringList(base.AddRange(values));
        }
        public new NonNullOrEmptyStringList Insert(string value, int insertAtIndex)
        {
            return ToNonNullOrEmptyStringList(base.Insert(value, insertAtIndex));
        }
        public new NonNullOrEmptyStringList Insert(IEnumerable<string> values, int insertAtIndex)
        {
            return ToNonNullOrEmptyStringList(base.Insert(values, insertAtIndex));
        }
        public new NonNullOrEmptyStringList Remove(string value)
        {
            return ToNonNullOrEmptyStringList(base.Remove(value));
        }
        public new NonNullOrEmptyStringList Remove(string value, IEqualityComparer<string> optionalComparer)
        {
            return ToNonNullOrEmptyStringList(base.Remove(value, optionalComparer));
        }
        public new NonNullOrEmptyStringList RemoveAt(int removeAtIndex)
        {
            return ToNonNullOrEmptyStringList(base.RemoveAt(removeAtIndex));
        }
        public new NonNullOrEmptyStringList RemoveRange(int removeAtIndex, int count)
        {
            return ToNonNullOrEmptyStringList(base.RemoveRange(removeAtIndex, count));
        }
        public new NonNullOrEmptyStringList Sort()
        {
            return ToNonNullOrEmptyStringList(base.Sort());
        }
        public new NonNullOrEmptyStringList Sort(Comparison<string> optionalComparison)
        {
            return ToNonNullOrEmptyStringList(base.Sort(optionalComparison));
        }
        public new NonNullOrEmptyStringList Sort(IComparer<string> optionalComparer)
        {
            return ToNonNullOrEmptyStringList(base.Sort(optionalComparer));
        }

        private static NonNullOrEmptyStringList ToNonNullOrEmptyStringList(ImmutableList<string> list)
        {
            if (list == null)
                throw new ArgumentNullException("list");

            return To<NonNullOrEmptyStringList>(
                list,
                tail => new NonNullOrEmptyStringList(tail)
            );
        }

        private class Validator : IValueValidator<string>
        {
            /// <summary>
            /// This will throw an exception for a value that does pass validation requirements
            /// </summary>
            public void EnsureValid(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                    throw new ArgumentException("Null/blank value specified");
            }
        }
    }
}
