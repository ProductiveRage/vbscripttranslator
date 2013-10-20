using System;

namespace CSharpWriter.Lists
{
    public static class ImmutableList_Extensions
    {
        /// <summary>
        /// It's common for the LINQ Any() extension method to be used with any enumerable sets but with the ImmutableList this can be a relatively expensive
        /// operation as it has to traverse all of its items in order to enable enueration - if all we want to know is if there are any elements in it then
        /// it's much cheaper to query the Count property. As this is more specific method than the LINQ extension method, this will be used in preference
        /// by the compiler.
        /// </summary>
        public static bool Any<T>(this ImmutableList<T> data)
        {
            if (data == null)
                throw new ArgumentNullException("data");

            return (data.Count > 0);
        }

        /// <summary>
        /// The LINQ Count() extension method enumerates over the entire set which can be a relatively expensive operation with the ImmutableList, much
        /// more expensive than querying the Count property directly, at least. As this is more specific method than the LINQ extension method, this
        /// will be used in preference by the compiler.
        /// </summary>
        public static int Count<T>(this ImmutableList<T> data)
        {
            if (data == null)
                throw new ArgumentNullException("data");

            return data.Count;
        }
    }
}
