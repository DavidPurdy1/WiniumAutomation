using System;

namespace WiniumTests.src {
    class TestData {
        public int TestsFailed { get; set; }
        public int TestsPassed { get; set; }
        public string CreatedBy { get; set; }
        public DateTime CreatedDate { get; set; }
        public string ImagePath { get; set; }
        public string ApplicationName { get; set; }
        public string ApplicationVersion { get; set; }
        public string TestName { get; set; }
        public int TestStatus { get; set; }
        public int TestRunId { get; set; }
    }
}
