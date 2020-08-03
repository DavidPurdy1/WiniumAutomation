using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using System.Configuration;
using System.Reflection;
using System.Threading;

namespace WiniumTests.src {
    /// <summary>
    /// Methods that happen at the beginning of a test
    /// </summary>
    public class IntactSetup {

        IWebElement window;
        readonly WiniumMethods m;
        string method = "";

        public IntactSetup(WiniumMethods m) {
            this.m = m;
        }
        /**Doesn't work
         * This is going to connect to a remote server with the name brought in by serverName, when used
         * Doesn't work with the changes that have been made, lots of these are readonly
         */
        private void ConnectToRemoteDesktop(AppiumOptions options,string serverName) {

            //method = MethodBase.GetCurrentMethod().Name;
            //Print(method, "Started");

            //options.ApplicationPath = ConfigurationManager.AppSettings.Get("RemoteDesktop");
            //driver = new WiniumDriver(driverPath, options);
            //m.sendKeysByName("Remote Desktop Connection", serverName);
            //m.clickByName("Connect");

            //print(method, "Finished");
        }
        /**THIS METHOD HAS TO BE RAN FIRST WITH TESTS, Logs into intact with admin login
         */
        public void Login() { //TODO: have to add connectToRemoteDesktop
            method = MethodBase.GetCurrentMethod().Name;
            Thread.Sleep(12000);
            m.SwitchWindowHandle();
            var window = m.Locate(By.Id("frmLogin"));
            m.Locate(By.Id("txtPassword"), window);
            Thread.Sleep(1000);
            window.SendKeys("admin");
            Thread.Sleep(1000);
            m.Click(By.Name("&Logon"), m.Locate(By.Id("frmLogin")));
            Thread.Sleep(1000);
            m.SwitchWindowHandle();
            m.handleList.Add(m.GetCurrentHandle());
        }
        public void Logout() {
            window = m.Locate(By.Name("&Intact"), m.Locate(By.Name("radMenu1")));
            m.Click(By.Name("Log Out"), window);
        }
        private void SetDatabaseInformation() {
            m.Click(By.Name("&Settings.."));
            m.SendKeys(By.Name(""), @"(local)\INTACT");
            m.SendKeys(By.Name(""), "{TAB}");
            m.SendKeys(By.Name(""), "{TAB}");
            m.SendKeys(By.Name(""), "{ENTER}");
        }
    }
}
