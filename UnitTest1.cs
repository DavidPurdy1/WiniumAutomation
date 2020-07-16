﻿using log4net;
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
        public TestContext TestContext { get; set; }
        #endregion

        #region
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
            Print(TestContext.TestName, "STARTED *********************************************");
        }
        [TestCleanup]
        public void TestCleanup() {
            if (TestContext.CurrentTestOutcome == UnitTestOutcome.Passed) {
                testsPassedNames.Add(TestContext.TestName);
                Print(method, "PASSED *****************************************");
            } else {
                string imagePath = user.OnFail(TestContext.TestName);
                testsFailedNames.Add(TestContext.TestName + " " + imagePath + " | ");
                Print(method, "FAILED *****************************************");
            }
        }
        [ClassCleanup]
        public static void Cleanup() {
            user.WriteFailFile(testsFailedNames, testsPassedNames);
            user.CloseDriver();
        }
        #endregion 

        #region
        [TestMethod]
        public void TEST1_LOGIN() {
            method = MethodBase.GetCurrentMethod().Name;
            user.LoginToIntact();
        }
        [TestMethod]
        public void TEST2_INZONE() { //change it so that it throws error in usermethod
            method = MethodBase.GetCurrentMethod().Name;
            if (!user.GetDocumentsFromInZone()) {
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
            user.LoginToIntact();
            user.CreateNewDefinition();
        }
        [TestMethod]
        public void TEST5_TYPES() {
            method = MethodBase.GetCurrentMethod().Name;
            user.LoginToIntact();
            user.CreateNewType();
        }
        [TestMethod]
        public void TEST6_DOCUMENTS() {
            method = MethodBase.GetCurrentMethod().Name;
            user.CreateDocument(1, true);
        }
        [TestMethod]
        public void TEST7_SEARCH() { //fix printings
            method = MethodBase.GetCurrentMethod().Name;
            user.Search("water");
        }
        [TestMethod]
        public void TEST8_RECOGNITION() {
            method = MethodBase.GetCurrentMethod().Name;
            user.TestRecognition("DEFAULT DOCUMENT OPTIONS", "DEFAULT DOCUMENT", "lorem");
        }
        [TestMethod]
        public void TEST9_UTIL() {
            method = MethodBase.GetCurrentMethod().Name;
            user.LoginToIntact();
            user.OpenUtil();
        }
        [TestMethod]
        public void TEST10_IPACK() {
            method = MethodBase.GetCurrentMethod().Name;
            user.AddToIPack();
        }
        [TestMethod]
        public void TEST11_LOGOUT() {
            method = MethodBase.GetCurrentMethod().Name;
            user.Logout();
        }
        #endregion

        private void Print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
    }
}