using System.Diagnostics;
using System.Collections.Generic;
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using log4net;
using System.Reflection;


namespace WiniumTests.src
{

    public class TestTemplate
    {
        public TestContext TestContext { get; set; }
        public static UserMethods user;
        string method = string.Empty;
        static readonly ILog debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
        static readonly List<string> testsFailedNames = new List<string>();
        static readonly List<string> testsPassedNames = new List<string>();
        static readonly List<string> testsInconclusiveNames = new List<string>();
        static readonly List<string> imagePaths = new List<string>();

        [ClassInitialize]
        public void ClassInitialize()
        {
            method = MethodBase.GetCurrentMethod().Name;
            throw new NotImplementedException();
        }
        public void TestInitialize()
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

        public void TestCleanup()
        {
            method = MethodBase.GetCurrentMethod().Name;
            if (TestContext.CurrentTestOutcome == UnitTestOutcome.Inconclusive)
            {
                testsInconclusiveNames.Add(TestContext.TestName);
                Print(method, "INCONCLUSIVE *********************************");
            }
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
        public void ClassCleanup()
        {
            user.Cleanup().WriteFailFile(testsFailedNames, testsPassedNames, testsInconclusiveNames, imagePaths);
            user.Cleanup().SendToDB();
        }
        private void Print(string method, string toPrint)
        {
            debugLog.Info(method + " " + toPrint);
        }
    }

}