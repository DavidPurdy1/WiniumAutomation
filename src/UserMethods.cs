using log4net;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Configuration;

namespace WiniumTests.src {
    /// <summary>
    /// Entry point for writing tests
    /// </summary>
    public class UserMethods {
        #region fields
        readonly AppiumOptions options;
        readonly WindowsDriver<WindowsElement> driver;
        readonly WiniumMethods m;
        readonly ILog debugLog;
        readonly Actions action;
        protected const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723/";
        #endregion

        #region Setup
        public UserMethods(ILog log) {

            debugLog = log;
            options = new AppiumOptions();
            options.AddAdditionalCapability("app", ConfigurationManager.AppSettings.Get("IntactPath"));
            driver = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), options);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);
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