using OpenQA.Selenium;
using System.Configuration;
using System.Reflection;
using System.Threading;

namespace WiniumTests.src {
    class IntactSetup{

        IWebElement window;
        readonly WiniumMethods m;

        public IntactSetup(WiniumMethods m) {
            this.m = m;
        } 
        /**Doesn't work
         * This is going to connect to a remote server with the name brought in by serverName, when used
         * Doesn't work with the changes that have been made, lots of these are readonly
         */
        public void ConnectToRemoteDesktop( string serverName) {

            //method = MethodBase.GetCurrentMethod().Name;
            //print(method, "Started");

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
            
            //both of these will most likely stay false all the time, but we can test either of them by changing value in app.config
            bool needToSetDB = ConfigurationManager.AppSettings.Get("setDataBase") == "true"; ;
            bool connectToRemote = ConfigurationManager.AppSettings.Get("connectToRemote") == "true";
            Thread.Sleep(10000);
            m.SendKeys(By.Name(""), "admin");
            if (!needToSetDB) {
                m.Click(By.Name("&Logon"));
            } else {
                setDatabaseInformation();
                m.Click(By.Name("&Logon"));
            }
            Thread.Sleep(2000);
        }
        public void Logout() {
            window = m.Locate(By.Name("&Intact"), m.Locate(By.Name("radMenu1")));
            m.Click(By.Name("Log Out"), window);
        }
        private void setDatabaseInformation() {
            m.Click(By.Name("&Settings.."));
            m.SendKeys(By.Name(""), @"(local)\INTACT");
            m.SendKeys(By.Name(""), "{TAB}");
            m.SendKeys(By.Name(""), "{TAB}");
            m.SendKeys(By.Name(""), "{ENTER}");
        }
    }
}
