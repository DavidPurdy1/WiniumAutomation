using System;
using System.Reflection;
using System.Threading;
using log4net;
using log4net.Config;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Winium;
using System.Windows.Forms;
using com.infodynamics.logging.ui;
using System.IO;
using Winium.Elements.Desktop.Extensions;
using System.Diagnostics;
using OpenQA.Selenium.Internal;
using System.Linq.Expressions;

namespace WiniumTests { //TODO: CHECK OUT RAM USAGE LARGE AMOUNT OF INSTANCES OF WINIUM WEBDRIVER RUNNING IN TASKMANAGER **driver.close() in every situation 
    [TestClass]
    public class IntactTest {
        WiniumDriver driver;
        DesktopOptions options = new DesktopOptions();
        string driverPath = @"C:\Users\i00018\Downloads\Winium.Desktop.Driver"; //TODO: add to be able to be used on other machines: edit i00018
        bool needToSetDB = false;
        static ILog debugLog = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType.FullName);
        string method;

        [TestMethod]
        public void Test() { //TODO: If key is pressed it can exit the testing
            //TODO: RIGHT NOW IT IS IMPORTANT THAT INTACT IS FULLSCREEN WHEN LAUNCHED FOR TESTING REASONS
            string currentMethod = MethodBase.GetCurrentMethod().Name;
            print(currentMethod, "Started");
            XmlConfigurator.Configure();

            killPreviousIntact();
            loginToIntact(options);
            getDocumentsFromInZone();


            print(currentMethod, "Finished*********************************************");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Thread.Sleep(10000);
            driver.Close();
        }

        //user custom methods **************************************************
        /**
         * This is going to connect to a remote server with the name brought in by serverName, when used
         */
        public void connectToRemoteDesktop(DesktopOptions options, string serverName) {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");

            options.ApplicationPath = @"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Accessories\Remote Desktop Connection.lnk";
            driver = new WiniumDriver(driverPath, options);
            Thread.Sleep(5000);
            sendKeysByName("Remote Desktop Connection", serverName);
            Thread.Sleep(2000);
            clickByName("Connect");
            Thread.Sleep(5000);

            print(method, "Finished");
        }
        /**
         * Logs into intact automatically with the admin login, THIS METHOD HAS TO BE RAN FIRST WITH THE DRIVER 
         */
        public void loginToIntact(DesktopOptions options) {
            method = MethodBase.GetCurrentMethod().Name;
            debugLog.Info(method + " Started");

            options.ApplicationPath = @"C:\Program Files (x86)\InfoDynamics, Inc\Intact 4\Intact.exe"; //TODO: check here if another instance of Intact open or not
            driver = new WiniumDriver(driverPath, options);
            Thread.Sleep(7000);
            sendKeysByName("", "admin");
            if (!needToSetDB) {
                clickByName("&Logon");
                Thread.Sleep(3000);
            } else {
                setDatabaseInformation();
                clickByName("&Logon");
                Thread.Sleep(3000);
            }
            debugLog.Info(method + " Finished");
        }
        /**
         * This is going to a specified amount of definitions with random name for each blank. TODO: RIGHT NOW IT IS IMPORTANT THAT INTACT IS FULLSCREEN WILL NOT WORK OTHERWISE
         */
        public void createNewDefinition(int numberOfDefinitions) {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");

            var window = driver.FindElement(By.Id("frmIntactMain"));

            window = window.FindElement(By.Name("radMenu1"));
            clickByNameInTree(window, "&Administration");

            window = window.FindElement(By.Name("&Administration"));
            clickByNameInTree(window, "Definitions");


            Thread.Sleep(5000);

            for (int i = 0; i <= numberOfDefinitions; i++) {
                var temp = new Random().Next().ToString();

                window = driver.FindElementById("frmIntactMain");

                window = window.FindElement(By.Id("frmRulesList"));

                window.FindElement(By.Id("btnAdd")).Click();

                window = window.FindElement(By.Name("Add Definition"));

                debugLog.Info("Definition name is " + "Test " + temp);

                foreach (IWebElement element in window.FindElements(By.Name(""))) {
                    if (element.Enabled == true) {
                        try { element.SendKeys("Test " + temp); } catch (Exception) { }
                    }
                }
                clickByNameInTree(window, "&Save");
            }
            clickByName("&Close");
            print(method, "Finished");
        }
        /**
         * Going to create a new type and will add random values for all of blanks. VERY IMPORTANT THAT THE INTACT WINDOW IS FULLSCREENED FOR RIGHT NOW 
         */
        public void createNewType(int numberOfTypes) {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");

            var window = driver.FindElement(By.Id("frmIntactMain"));

            window = window.FindElement(By.Name("radMenu1"));

            clickByNameInTree(window, "&Administration");

            window = window.FindElement(By.Name("&Administration")); //add the setting of the window in the tree method and rename it
            clickByNameInTree(window, "Types");


            Thread.Sleep(2000);
            for (int i = 0; i < numberOfTypes; i++) {
                var temp = new Random().Next().ToString();

                window = driver.FindElement(By.Id("frmIntactMain"));
                Thread.Sleep(2000);

                window = window.FindElement(By.Id("frmAdminTypes"));
                window.FindElement(By.Id("rbtnAdd")).Click();

                window = window.FindElement(By.Id("frmAdminTypesInfo"));

                foreach (IWebElement element in window.FindElements(By.Name(""))) {
                    if (element.Enabled == true) {
                        try { element.SendKeys("Test " + temp); } catch (Exception) { }
                    }
                }
                clickByName("&OK");
            }
            clickByName("&Close");
            print(method, "Finished");
        }
        //UNTESTED
        public void setDatabaseInformation() {
            method = MethodBase.GetCurrentMethod().Name;
            debugLog.Info(method + " Started");
            clickByName("&Settings..");
            sendKeysByName("", @"(local)\INTACT");
            sendKeysByName("", "{TAB}");
            sendKeysByName("", "{TAB}");
            sendKeysByName("", "{ENTER}");

            debugLog.Info(method + " Finished");
        }
        public void openOrganizer() {
            method = MethodBase.GetCurrentMethod().Name;
            debugLog.Info(method + " Started");
            clickByName("Organizer");
            Thread.Sleep(2000);
            debugLog.Info(method + " Finished");
        }
        /**
         * Takes in isPDF telling whether or not it is a PDF or TIF and then numOfDocs for how many to add.
         * Going to take a random document from some preset path and with that it is going to add that document either pdf or not to a random definition and type with random metadata.
         */
        public void createDocument(bool isPDF, int numOfDocs) {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            var window = driver.FindElement(By.Id("frmIntactMain"));
            window = window.FindElement(By.Id("radPanelBar1"));
            window = window.FindElement(By.Id("pageIntact"));
            window = window.FindElement(By.Id("lstIntact"));
            clickByNameInTree(window, "Add Document");

            window = driver.FindElement(By.Id("frmDocument"));
            window = window.FindElement(By.Id("radCommandBar1"));

            clickByNameInTree(window, "btnAdd");
            print(method, "IT REACHED HERE +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
            //window = driver.FindElement(By.Name("Desktop 1"));
            //window = window.FindElement(By.Name("DropDown"));
            clickByNameInTree(window, "Add Native File...");
            for (int i = 0; i < numOfDocs; i++) {
                //TODO: Add 10 pdf and tifs to the folder where it adds documents from
                if (isPDF) {
                    var rand = new Random();
                    sendKeysByName("File Name:", "PDF" + rand.Next(10).ToString());
                } else {
                    var rand = new Random();
                    sendKeysByName("File Name:", "TIF" + rand.Next(10).ToString());
                }
            }
            print(method, "Finished");
        }
        /**
         * Collects documents by all definition from InZone
         */
        public void getDocumentsFromInZone() {
            var window = driver.FindElement(By.Id("frmIntactMain"));
            window = findById(window, "radMenu1");
            clickByNameInTree(window, "&Intact");
            window = findByName(window, "&Intact");
            clickByNameInTree(window, "InZone");
            window = driver.FindElement(By.Id("frmInZoneMain"));
            clickByIdInTree(window, "btnCollectScan");
        }
        //UNTESTED: ALSO NEED TO ADD COMMIT TO GETDOCUMENTS
        /**
         * Method copies over files from one directory to another: Each time before InZone collects this is going to put files in the collector folder
         * Verify that the startPath always has files in it and those files shouldn't be removed from this folder when collected
         */
        public void addDocsToCollector() {
            string startPath = @"";
            string endPath = @"";
            if (Directory.Exists(startPath)) {
                foreach(string s in Directory.GetFiles(startPath)) {
                    var fileName = Path.GetFileName(s);
                    File.Copy(fileName, Path.Combine(endPath, fileName), true); 
                }
            } else {
                print(method, "Starting path doesn't exist");
            }
        }

        //winium custom methods ************ TODO: Whenever it throws an error use the driver to take a screenshot

        /**
         * Takes in the string of the element name that can be found from inspect.exe from Windows SDK
         * Will throw exception if element name is not found. Always use Name if the element is named
         */
        private void clickByName(string elementName) { // checks if the element is present multiple times review later to make sure that it is working efficiently
            IWebElement element = null;
            method = MethodBase.GetCurrentMethod().Name;
            for (int i = 1; i <= 5; i++) {
                try {
                    element = driver.FindElementByName(elementName);

                    if (element != null) {
                        break;
                    }
                } catch (NoSuchElementException e) {
                    debugLog.Info(method + " ", e);
                    Thread.Sleep(1000);
                }
            }
            for (int i = 1; i <= 5; i++) { // checks if the element is enabled, only happens if something has a name, but is not clickable
                try {
                    if (element.Enabled == true) {
                        debugLog.Info(method + " " + elementName + " " + "Clicked");
                        element.Click();
                        Thread.Sleep(2000);
                        break;
                    }
                } catch (NoSuchElementException e) {
                    debugLog.Info(method + " ", e);
                }
            }
        }

        /**
     * Takes in the string of the element name that can be found from inspect.exe from Windows SDK and a input string to send to the field 
     * Will throw exception if element name is not found or is a field that cannot be edited. Always use Name if the element is named
     */
        private void sendKeysByName(string elementName, string input) { //TODO: fix the method because it needs to stop throwing errors on every time
            IWebElement element = null;
            method = MethodBase.GetCurrentMethod().Name;
            for (int i = 1; i <= 5; i++) {
                try {
                    element = driver.FindElementByName(elementName); break;
                } catch (NoSuchElementException e) {
                    debugLog.Info(method + " ", e);
                    Thread.Sleep(1000);
                }
            }
            for (int i = 1; i <= 5; i++) {
                try {
                    if (element.Enabled == true) {
                        element.SendKeys(input);
                        debugLog.Info(method + " " + elementName + " " + "Sent keys");
                        Thread.Sleep(2000);
                        break;
                    }
                } catch (NoSuchElementException e) {
                    Thread.Sleep(1000);
                    debugLog.Info(method + " ", e); ;
                }
            }
        }

        /**
         * Takes in the string of the element Id that can be found from inspect.exe from Windows SDK and a input string to send to the field
         * Will throw exception if element name is not found or is a field that cannot be edited.
         */
        private void sendKeysById(string elementId, string input) { //TODO: fix the method because it needs to stop throwing errors on every time
            IWebElement element = null;
            method = MethodBase.GetCurrentMethod().Name;
            for (int i = 1; i <= 5; i++) {
                try {
                    element = driver.FindElementById(elementId); break;
                } catch (Exception e) {
                    debugLog.Info(method + " ", e);
                    Thread.Sleep(1000);
                }
            }
            for (int i = 1; i <= 10; i++) {
                try {
                    if (element.Enabled == true) {
                        element.SendKeys(input); break;
                    }
                } catch (Exception e) {
                    debugLog.Info(method + e.StackTrace);
                    debugLog.Info(method + " ", e); ;
                }
            }
        }

        /**
         * This method is going to take something in the pre-existing element blank and add the input string to the end of it without a space
         */
        private void addInputToEntryByClass(string elementName, string input) {
            IWebElement element = null;
            method = MethodBase.GetCurrentMethod().Name;
            for (int i = 1; i <= 5; i++) {
                try {
                    element = driver.FindElementByName(elementName); break;
                } catch (Exception e) {
                    debugLog.Info(method + " ", e);
                    Thread.Sleep(1000);
                }
            }
            for (int i = 1; i <= 5; i++) {
                try {
                    if (element.Enabled == true) {
                        string value = driver.FindElementByClassName(elementName).Text;
                        driver.FindElementByClassName(elementName).SendKeys(value + input);
                        break;
                    }
                } catch (Exception e) {
                    Thread.Sleep(1000);
                    debugLog.Info(method + " ", e);
                }
            }
        }
        /**
         * Got rid of the for loops for checking if the elements are enabled
         */
        private void clickByNameInTree(IWebElement parent, string elementName) {
            var element = findByName(parent, elementName);
            try {
                if (element.Enabled == true) {
                    element.Click();
                    debugLog.Info(method + " " + elementName + " Clicked");
                    Thread.Sleep(2000);
                }
            } catch (NoSuchElementException e) {
                printError(method, null, e);
            }
        }
        private void clickByIdInTree(IWebElement parent, string elementName) {
            var element = findById(parent, elementName);
            try {
                if (element.Enabled == true) {
                    element.Click();
                    debugLog.Info(method + " " + elementName + " Clicked");
                    Thread.Sleep(2000);
                }
            } catch (NoSuchElementException e) {
                printError(method, null, e);
            }
        }
        private IWebElement findById(IWebElement parent, string elementName) {
            IWebElement element = null;
            method = MethodBase.GetCurrentMethod().Name;

            for (int i = 1; i <= 5; i++) {
                try {
                    element = parent.FindElement(By.Id(elementName));
                    break;
                } catch (NoSuchElementException e) {
                    debugLog.Info(method + " ", e);
                    Thread.Sleep(1000);
                }
            }
            return element;
        }
        private IWebElement findByName(IWebElement parent, string elementName) {
            IWebElement element = null;
            method = MethodBase.GetCurrentMethod().Name;
            for (int i = 1; i <= 5; i++) {
                try {
                    element = parent.FindElement(By.Name(elementName));
                    break;
                } catch (NoSuchElementException e) {
                    debugLog.Info(method + " ", e);
                    Thread.Sleep(1000);
                }
            }
            return element;
        }
        /**
         * Checks to see if there is already an instance of Intact running, if so Kills it so a new intact test can run
         * TODO: SEE IF THERE IS A WAY TO DISPOSE OF THE APPLICATION WITHOUT KILLING
         */
        private void killPreviousIntact() {
            foreach (Process app in Process.GetProcesses()) {
                if (app.ProcessName.Equals("Intact")) {
                    app.Kill();
                    print(method, "Past Intact Killed");
                }
            }
        }

        //Want this to cycle through all processes like kill method, 
        //but then cast to a window and get the windowState, then have it set the window max based off that 
        private bool isMaximized() {
            //FormWindowState
            return false;

        }

        //debugging methods

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