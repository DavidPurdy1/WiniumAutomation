using System;
using System.Reflection;
using log4net;
using log4net.Config;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Windows.Forms;

namespace WiniumTests { //TODO: CHECK OUT MEMORY USAGE LARGE AMOUNT OF INSTANCES OF WINIUM WEBDRIVER RUNNING IN TASKMANAGER **driver.close() in every situation 
    [TestClass]
    public class IntactTest {
        static ILog debugLog;
        string method;
        UserMethods user;

        public IntactTest() {
            debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
            user = new UserMethods(debugLog);
            XmlConfigurator.Configure();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
        }

        [TestMethod]
        public void test() { //TODO: If key is pressed it can exit the testing    //TODO: RIGHT NOW IT IS IMPORTANT THAT INTACT IS FULLSCREEN WHEN LAUNCHED FOR TESTING REASONS                
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            user.loginToIntact();
            user.createDocument();



            print(method, "Finished*********************************************");
        }

        [TestMethod]
        public void InZonetest() {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            user.loginToIntact();
            if (!user.getDocumentsFromInZone()) {
                Assert.Fail();
            }
            print(method, "Finished*********************************************");
        }

        [TestMethod]
        public void BatchReviewtest() {           
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            user.loginToIntact();
            user.BatchReview();
            print(method, "Finished*********************************************");
        }

        [TestMethod]
        public void testAll() {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started Last Test");
            test();
            InZonetest();
            BatchReviewtest();
            
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