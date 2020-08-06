using log4net;
using log4net.Config;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using WiniumTests.src;

namespace WiniumTests {
    /// <summary>
    /// Login to Intact, Create a Document, Search, Recognize, InZone, Batch Review
    /// Time: About 2 mins
    /// </summary>
    [TestClass]
    public class FullTest {
        #region Test Fields
        static readonly ILog debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
        public string method;
        public static UserMethods user;
        static readonly List<string> testsFailedNames = new List<string>();
        static readonly List<string> testsPassedNames = new List<string>();
        static readonly List<string> imagePaths = new List<string>();
        public TestContext TestContext { get; set; }
        #endregion

        #region Test Attributes
        [ClassInitialize]
        public static void ClassInit(TestContext testContext) {
            user = new UserMethods(debugLog);
            XmlConfigurator.Configure();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
        }
        [TestInitialize]
        public void TestInit() {
            method = MethodBase.GetCurrentMethod().Name;
            foreach (Process app in Process.GetProcesses()) {
                if (app.ProcessName.Equals("Intact")) {
                    app.Kill();
                    Print(method, "Previous Intact Killed");
                }
            }
            user = new UserMethods(debugLog);
            Print(TestContext.TestName, "STARTED *********************************************");
        }
        [TestCleanup]
        public void TestCleanup() {
            if (TestContext.CurrentTestOutcome == UnitTestOutcome.Passed) {
                testsPassedNames.Add(TestContext.TestName);
                Print(method, "PASSED *****************************************");
            } else {
                imagePaths.Add(user.Cleanup().OnFail(TestContext.TestName) + ".PNG");
                testsFailedNames.Add(TestContext.TestName);
                Print(method, "FAILED *****************************************");
            }
            user.Cleanup().CloseDriver();
        }
        [ClassCleanup]
        public static void Cleanup() {
            user.Cleanup().WriteFailFile(testsFailedNames, testsPassedNames, imagePaths);
            user.Cleanup().SendToDB();
        }
        #endregion

        #region Print
        private void Print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
        #endregion

        [TestMethod]
        public void TEST1_2_INZONE() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.DocumentCollect().InZone();
        }
        [TestMethod]
        public void TEST1_3_BATCHREVIEW() { //Batch review runs slow      
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.DocumentCollect().BatchReview();
        }
        [TestMethod]
        public void TEST1_6_DOCUMENTS() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.Create().SimpleCreateDocument();
        }
        [TestMethod]
        public void TEST1_7_SEARCH() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.SearchRecognize().Search("Default");
        }
        [TestMethod]
        public void TEST1_8_RECOGNITION() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.SearchRecognize().Recognition("DEFAULT DOCUMENT OPTIONS", "DEFAULT DOCUMENT", "lorem");
        }

    }
}