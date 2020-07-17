

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace WiniumTests {
    public abstract class TestProperties : TestContext {
        public bool FailToken { get; set; }
    }
}
