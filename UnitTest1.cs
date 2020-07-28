using log4net;
using log4net.Config;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace WiniumTests {
    [TestClass]
    public class IntactTest {
        #region
        static readonly ILog debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
        public string method;
        public static UserMethods user;
        static List<string> testsFailedNames = new List<string>();
        static List<string> testsPassedNames = new List<string>();
        static List<string> imagePaths = new List<string>();
        public TestContext TestContext { get; set; }
        #endregion

        #region
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
                imagePaths.Add(user.OnFail(TestContext.TestName) + ".PNG");
                testsFailedNames.Add(TestContext.TestName);
                Print(method, "FAILED *****************************************");
            }
            user.CloseDriver();
        }
        [ClassCleanup]
        public static void Cleanup() {
            user.WriteFailFile(testsFailedNames, testsPassedNames, imagePaths);
            user.SendToDB(); 
        }
        #endregion 

        #region
        [TestMethod]
        public void TEST1_1_LOGIN() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
        }
        [TestMethod]
        public void TEST1_2_INZONE() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.InZone();
        }
        [TestMethod]
        public void TEST1_3_BATCHREVIEW() { //Batch review runs slow      
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.BatchReview();
        }
        [TestMethod]
        public void TEST1_4_DEFINITIONS() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.CreateNewDefinition();
        }
        [TestMethod]
        public void TEST1_5_TYPES() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.CreateNewType();
        }
        [TestMethod]
        public void TEST1_6_DOCUMENTS() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.CreateDocument(1, true);
        }
        [TestMethod]
        public void TEST1_7_SEARCH() { 
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.Search("water");
        }
        [TestMethod]
        public void TEST1_8_RECOGNITION() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.TestRecognition("DEFAULT DOCUMENT OPTIONS", "DEFAULT DOCUMENT", "lorem");
        }
        [TestMethod]
        public void TEST2_1_IPACK() { //unfinished
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.AddToIPack();
        }
        [TestMethod]
        public void TEST2_2_LOGOUT() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.Logout();
        }
        [TestMethod]
        public void TEST2_3_AUDITTRAIL() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Login();
            user.AuditTrail();
        }
        //[TestMethod]
        //public void TEST1_9_UTIL() { //unfinished 
        //    method = MethodBase.GetCurrentMethod().Name;
        //    user.Login();
        //    user.OpenUtil();
        //}
        //[TestMethod]
        //public void TEST2_4_PORTFOLIO() {
        //    method = MethodBase.GetCurrentMethod().Name;
        //    user.Login();
        //    user.Portfolio(new string[] {"time", "is", "10:48" });
        //}

        #region RUN ALL TESTS HERE *******
        /**
         * This test is going to go straight through all testing without closing Intact. 
         * Inaccurate compared to testing each one in a playlist, but faster. 
         * Needs to be updated each time a new test is created.
         */
        [TestMethod]
        public void TEST9_9_ALL() {  //do something with a global boolean to decide whether or not you have to login. Then you can call the tests inside of another test.
            method = MethodBase.GetCurrentMethod().Name;
            //user.Login();
            //user.CreateNewType();
            //user.CreateNewDefinition();
            //user.CreateDocument();
            //user.InZone();
            //user.BatchReview();
            //user.Search("test");
            //user.TestRecognition("DEFAULT DEF", "DEFAULT DEFINITION TEST", "test");
            //user.Portfolio();
            //user.AuditTrail();
            //user.AddToIPack();
            //user.OpenUtil();
            //user.Logout();
        }
        #endregion

        #endregion

        private void Print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
    }
}