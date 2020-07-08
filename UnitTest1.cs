using log4net;
using log4net.Config;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace WiniumTests {
    [TestClass]
    public class IntactTest {
        static readonly ILog debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
        public string method;
        public static UserMethods user;
        private static TestContext testContext;
        static List<string> testsFailedNames = new List<string>();
        static List<string> testsPassedNames = new List<string>();

        [AssemblyInitialize]
        public static void TestingInit(TestContext _testContext) { //TODO: something having to do with the testContext is wrong
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
        public void TestCleanup() { //if the test is failed then it has to close out of the current window and return to the intact main look at onFail in usermethods
            if (testContext.CurrentTestOutcome == UnitTestOutcome.Passed) {
                testsPassedNames.Add(testContext.TestName);
                print(method, "PASSED *****************************************");
            } else {
                testsFailedNames.Add(testContext.TestName);
                user.failLog(testContext.TestName);
                user.onFail();
                print(method, "FAILED *****************************************");
            }
        }
        [ClassCleanup]
        public static void Cleanup() { //add test report at the end, either excel over time or something to record test data.
            user.writeFailFile(testsFailedNames, testsPassedNames);
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
                throw new AssertFailedException(method + " InZone did not recognize the definition correctly");
            }
        }
        [TestMethod]
        public void TEST3_BATCHREVIEW() { //Batch review runs slow      
            method = MethodBase.GetCurrentMethod().Name;
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
            user.loginToIntact();
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