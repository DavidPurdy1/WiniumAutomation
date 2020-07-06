using log4net;
using log4net.Config;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Remote;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace WiniumTests { //TODO: CHECK OUT MEMORY USAGE LARGE AMOUNT OF INSTANCES OF WINIUM WEBDRIVER RUNNING IN TASKMANAGER **driver.close() in every situation 
    [TestClass]
    public class IntactTest {
         ILog debugLog;
        public string method;
        public UserMethods user;
        
        //[AssemblyInitialize]
       // public void TestingInit() {
          //  debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
           // user = new UserMethods(debugLog);
           // XmlConfigurator.Configure();
          //  Application.EnableVisualStyles();
           // Application.SetCompatibleTextRenderingDefault(false);
     //   }


      //  [ClassCleanup]
     //   public void Cleanup() {
       //     user.closeDriver();
      //  }
    //
        public IntactTest() {
            killIntact();
            debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
            user = new UserMethods(debugLog);
            XmlConfigurator.Configure();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
       }

        [TestMethod]
        public void TEST1_LOGIN() { //TODO: If key is pressed it can exit the testing    //TODO: RIGHT NOW IT IS IMPORTANT THAT INTACT IS FULLSCREEN WHEN LAUNCHED FOR TESTING REASONS                
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            user.loginToIntact();
            print(method, "Finished*********************************************");
        }

        [TestMethod]
        public void TEST2_INZONE() {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            user.loginToIntact();

            if (!user.getDocumentsFromInZone()) {
                Assert.Fail();
            }

            user.closeDriver();
            print(method, "Finished*********************************************");
        }

        //NEED TO FIX
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

        public void killIntact() {
            foreach (Process app in Process.GetProcesses()) {
                if (app.ProcessName.Equals("Intact")) {
                    app.Kill();
                    print(method, "Past Intact Killed");
                }
            }
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