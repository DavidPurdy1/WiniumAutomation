using System;
using System.Runtime.Serialization;

namespace WiniumTests {
    [Serializable]
    internal class ElementNullException : Exception {
        public ElementNullException() {
        }

        public ElementNullException(string message) : base(message) {
        }

        public ElementNullException(string message, Exception innerException) : base(message, innerException) {
        }

        protected ElementNullException(SerializationInfo info, StreamingContext context) : base(info, context) {
        }
    }
}