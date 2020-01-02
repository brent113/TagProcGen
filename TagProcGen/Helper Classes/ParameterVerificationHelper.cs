using System;

namespace TagProcGen
{
    /// <summary>
    /// Extension method for parameter null checking
    /// </summary>
    public static class ParameterVerificationHelper
    {
        /// <summary>
        /// Throw if parameter is null
        /// </summary>
        public static void ThrowIfNull<T>(this T o, string parameterName)
        {
            if (o == null)
            {
                throw new ArgumentNullException(parameterName);
            }
        }
    }
}
