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
        static List<string> testsFailedNames = new List<string>();
        static List<string> testsPassedNames = new List<string>();
        public TestContext TestContext { get; set; }

        [AssemblyInitialize]
        public static void TestingInit(TestContext testContext) {
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
            print(TestContext.TestName, "STARTED *********************************************");
        }
        [TestCleanup]
        public void TestCleanup() { 
            if (TestContext.CurrentTestOutcome == UnitTestOutcome.Passed) {
                testsPassedNames.Add(TestContext.TestName);
                print(method, "PASSED *****************************************");
            } else {
                string imagePath = user.onFail(TestContext.TestName); 
                testsFailedNames.Add(TestContext.TestName + " " + imagePath + " | ");
                print(method, "FAILED *****************************************");
            }
        }
        [ClassCleanup]
        public static void Cleanup() { 
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
        public void TEST7_SEARCH() { //fix printings
            method = MethodBase.GetCurrentMethod().Name;
            user.loginToIntact();
            Assert.IsTrue(user.search("InZone"), " Search not found" ); 
        }
        [TestMethod]
        public void TEST8_RECOGNITION() {
            method = MethodBase.GetCurrentMethod().Name;
            user.loginToIntact();
            user.testRecognition("DEFAULT DEF" , "DEFAULT DEFINITION TEST", "DEFAULT DEFINITION TEST");
        }
        [TestMethod]
        public void TEST9_TEST() {
            method = MethodBase.GetCurrentMethod().Name;
            user.loginToIntact();
            user.OpenUtil();
        }

        private void print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
    }
}