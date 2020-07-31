using log4net;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Winium;
using System;
using System.Configuration;

namespace WiniumTests.src {
    /// <summary>
    /// Entry point for writing tests
    /// </summary>
    public class UserMethods {
        #region fields
        readonly DesktopOptions options = new DesktopOptions();
        readonly WiniumDriver driver;
        readonly WiniumMethods m;
        readonly ILog debugLog;
        readonly Actions action;
        #endregion

        #region Setup
        public UserMethods(ILog log) {
            debugLog = log;
            options.ApplicationPath = ConfigurationManager.AppSettings.Get("IntactPath");
            driver = new WiniumDriver(ConfigurationManager.AppSettings.Get("DriverPath"), options);
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            action = new Actions(driver);
            m = new WiniumMethods(driver, debugLog);
        }
        #endregion

        public IntactSetup Setup() {
            return new IntactSetup(m); 
        }
        public Create Create() {
            return new Create(m, action, debugLog); 
        }
        public Cleanup Cleanup() {
            return new Cleanup(m, debugLog); 
        }
        public DocumentCollect DocumentCollect() {
            return new DocumentCollect(m, action, debugLog); 
        }
        public SearchRecognize SearchRecognize() {
            return new SearchRecognize(m, action, debugLog);
        }
        public Misc Misc() {
            return new Misc(m, action,debugLog); 
        }

        public void Print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
    }
}