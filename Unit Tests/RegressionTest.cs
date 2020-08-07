using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using log4net;
using System.Reflection;
using WiniumTests.src;
using System.Diagnostics;
using log4net.Config;
using System.Windows.Forms;

namespace WiniumTests.Unit_Tests
{
    /// <summary>
    /// All Intact Tests with all elements tested
    /// Time: 
    /// </summary>
    [TestClass]
    public class RegressionTest
    {
        #region Test Fields
        static readonly ILog debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
        public string method;
        public static UserMethods user;
        static readonly List<string> testsFailedNames = new List<string>();
        static readonly List<string> testsPassedNames = new List<string>();
        static readonly List<string> testsInconclusiveNames = new List<string>();
        static readonly List<string> imagePaths = new List<string>();
        public TestContext TestContext { get; set; }
        #endregion

        #region Test Attributes
        [ClassInitialize]
        public static void ClassInit(TestContext testContext)
        {
            user = new UserMethods(debugLog);
            XmlConfigurator.Configure();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
        }
        [TestInitialize]
        public void TestInit()
        {
            method = MethodBase.GetCurrentMethod().Name;
            foreach (Process app in Process.GetProcesses())
            {
                if (app.ProcessName.Equals("Intact"))
                {
                    app.Kill();
                    Print(method, "Previous Intact Killed");
                }
            }
            user = new UserMethods(debugLog);
            Print(TestContext.TestName, "STARTED *********************************************");
        }
        [TestCleanup]
        public void TestCleanup()
        {
            if (TestContext.CurrentTestOutcome == UnitTestOutcome.Passed)
            {
                testsPassedNames.Add(TestContext.TestName);
                Print(method, "PASSED *****************************************");
            }
            else
            {
                imagePaths.Add(user.Cleanup().OnFail(TestContext.TestName) + ".PNG");
                testsFailedNames.Add(TestContext.TestName);
                Print(method, "FAILED *****************************************");
            }
            user.Cleanup().CloseDriver();
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            user.Cleanup().WriteFailFile(testsFailedNames, testsPassedNames, testsInconclusiveNames, imagePaths);
            user.Cleanup().SendToDB();
        }
        #endregion

        #region Print
        private void Print(string method, string toPrint)
        {
            debugLog.Info(method + " " + toPrint);
        }
        #endregion

        #region
        [TestMethod]
        public void TEST1_1_LOGIN()
        {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
        }
        [TestMethod]
        public void TEST1_2_INZONE()
        {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.DocumentCollect().InZone();
        }
        [TestMethod]
        public void TEST1_3_BATCHREVIEW()
        { //Batch review runs slow      
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.DocumentCollect().BatchReview();
        }
        [TestMethod]
        public void TEST1_4_DEFINITIONS()
        {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.Create().CreateNewDefinition();
        }
        [TestMethod]
        public void TEST1_5_TYPES()
        {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.Create().CreateNewType();
        }
        [TestMethod]
        public void TEST1_6_DOCUMENTS()
        {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.Create().CreateDocument(1, true);
        }
        [TestMethod]
        public void TEST1_7_SEARCH()
        {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.SearchRecognize().Search("Default");
        }
        [TestMethod]
        public void TEST1_8_RECOGNITION()
        {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.SearchRecognize().Recognition("DEFAULT DOCUMENT OPTIONS", "DEFAULT DOCUMENT", "lorem");
        }
        [TestMethod]
        public void TEST2_1_IPACK()
        { //unfinished
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.Misc().AddToIPack();
        }
        [TestMethod]
        public void TEST2_2_LOGOUT()
        {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.Setup().Logout();
        }
        [TestMethod]
        public void TEST2_3_AUDITTRAIL()
        {
            method = MethodBase.GetCurrentMethod().Name;
            user.Setup().Login();
            user.Misc().AuditTrail();
        }
        #endregion
    }
}
