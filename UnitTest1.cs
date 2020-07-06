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

        //public TestContext Context { get; set; }

        [AssemblyInitialize]
        public static void TestingInit(TestContext Context) {
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

        [ClassCleanup]
        public static void Cleanup() {
            user.closeDriver();
        }
        //Do something on test pass or failing, need to figure out how to have it determine if it passes or fails

        //[TestCleanup]
        //public void TestCleanup() {
        //    if (this.Context.CurrentTestOutcome == UnitTestOutcome.Failed) {
        //        print(method, "FAILED *****************************************");
        //    } else {
        //        print(method, "PASSED *****************************************");
        //    }
        //}

        [TestMethod]
        public void TEST1_LOGIN() {               
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            user.loginToIntact();
            print(method, "Finished*********************************************");
        }

        [TestMethod]
        public void TEST2_INZONE() {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");

            if (!user.getDocumentsFromInZone()) {
                Assert.Fail();
            }
            print(method, "Finished*********************************************");
        }

        [TestMethod]
        public void TEST3_BATCHREVIEW() {           
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            user.BatchReview();
            print(method, "Finished*********************************************");
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
            print(method, "Started Last Test");

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