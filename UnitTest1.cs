using log4net;
using log4net.Config;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace WiniumTests {
    [TestClass]
    public class IntactTest {
        static ILog debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
        public string method;
        public static UserMethods user;
        private static TestContext testContext;

        [AssemblyInitialize]
        public static void TestingInit(TestContext _testContext) {
            testContext = _testContext;
            foreach (Process app in Process.GetProcesses()) {
                if (app.ProcessName.Equals("Intact")) {
                    app.Kill();
                }
            }
            user = new UserMethods(debugLog);
            XmlConfigurator.Configure();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
        }
        [TestInitialize]
        public void TestInit() {
            print(testContext.TestName, "STARTED *********************************************");
        }
        [TestCleanup]
        public void TestCleanup() { //if the test is failed then it has to close out of the current window and return to the intact main 
            if (testContext.CurrentTestOutcome == UnitTestOutcome.Failed) {
                user.onFail();
                print(method, "FAILED *****************************************");
            } else {
                print(method, "PASSED *****************************************");
            }
        }
        [ClassCleanup]
        public static void Cleanup() {
            user.closeDriver();
        }


        [TestMethod]
        public void TEST1_LOGIN() {               
            method = MethodBase.GetCurrentMethod().Name;  
            user.loginToIntact();
        }
        [TestMethod]
        public void TEST2_INZONE() {
            method = MethodBase.GetCurrentMethod().Name;
            if (!user.getDocumentsFromInZone()) {
                Assert.Fail();
            }
        }
        [TestMethod]
        public void TEST3_BATCHREVIEW() { //Batch review runs slow      
            method = MethodBase.GetCurrentMethod().Name;
            user.loginToIntact();
            user.BatchReview();
        }
        [TestMethod]
        public void TEST4_DEFINITIONS() {
            method = MethodBase.GetCurrentMethod().Name;
            user.createNewDefinition();
        }
        [TestMethod]
        public void TEST5_TYPES() {
            method = MethodBase.GetCurrentMethod().Name;
            user.createNewType();
        }
        [TestMethod]
        public void TEST6_DOCUMENTS() {
            method = MethodBase.GetCurrentMethod().Name;
            user.createDocument();
        }


        [TestMethod]
        public void testAll() {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started Last Test");
            TEST1_LOGIN();
            TEST2_INZONE();
            TEST3_BATCHREVIEW();
            
            print(method, "All TESTS HAVE PASSED*********************************************");
        }
        [TestMethod]
        public void testsFailed() {
            method = MethodBase.GetCurrentMethod().Name;
        }

        /**
         * Prints to the log found in the temp folder
         */
        public void print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
        public void printError(string method, string toPrint, Exception e) {
            debugLog.Info(method + " " + e + " " + toPrint);
        }
    }
}