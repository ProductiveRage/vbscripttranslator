namespace CSharpWriter.Lists
{
    public interface IValueValidator<T>
    {
        /// <summary>
        /// This will throw an exception for a value that does pass validation requirements
        /// </summary>
        void EnsureValid(T value);
    }
}
