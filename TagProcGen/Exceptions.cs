using System;

namespace TagProcGen
{
    /// <summary>
    /// Exception for problems with tag generation.
    /// </summary>
    [Serializable]
    public class TagGenerationException : Exception
    {
        /// <summary>Default exception constructor.</summary>
        public TagGenerationException() { }
        /// <summary>Exception constructor with message.</summary>
        /// <param name="message">Message to throw exception with.</param>
        public TagGenerationException(string message) : base(message) { }
        /// <summary>Exception constructor with message.</summary>
        /// <param name="message">Message to throw exception with.</param>
        /// <param name="inner">Inner exception to include.</param>
        public TagGenerationException(string message, Exception inner) : base(message, inner) { }
        /// <summary>Serializable constructor.</summary>
        protected TagGenerationException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
