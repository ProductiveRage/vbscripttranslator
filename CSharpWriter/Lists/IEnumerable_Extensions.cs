using System.Collections.Generic;

namespace CSharpWriter.Lists
{
    public static class IEnumerable_Extensions
    {
        public static ImmutableList<T> ToImmutableList<T>(this IEnumerable<T> data)
        {
            return new ImmutableList<T>(data);
        }

        /// <summary>
        /// valueValidator is optional (may be null)
        /// </summary>
        public static ImmutableList<T> ToImmutableList<T>(this IEnumerable<T> data, IValueValidator<T> valueValidator)
        {
            return new ImmutableList<T>(data, valueValidator);
        }

        /// <summary>
        /// This will throw an exception if any of the values are null
        /// </summary>
        public static NonNullImmutableList<T> ToNonNullImmutableList<T>(this IEnumerable<T> data) where T : class
        {
            return new NonNullImmutableList<T>(data);
        }

        /// <summary>
        /// This will throw an exception if any of the values are null, valueValidator is optional (may be null)
        /// </summary>
        public static NonNullImmutableList<T> ToNonNullImmutableList<T>(this IEnumerable<T> data, IValueValidator<T> valueValidator) where T : class
        {
            return new NonNullImmutableList<T>(data, valueValidator);
        }
    }
}
