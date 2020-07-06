using log4net;
using OpenQA.Selenium;
using OpenQA.Selenium.Winium;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;

namespace WiniumTests {
    public class WiniumMethods { //Adding note for git commit test      
        string method;
        WiniumDriver driver;
        ILog debugLog;

        public WiniumMethods(WiniumDriver driver, ILog log) {
            this.driver = driver;
            debugLog = log;
        }

        /**
         * Takes in the string of the element name that can be found from inspect.exe from Windows SDK
         * Will throw exception if element name is not found. Always use Name if the element is named
         */
        public void clickByName(string elementName) { // checks if the element is present multiple times review later to make sure that it is working efficiently
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
        public void sendKeysByName(string elementName, string input) { //TODO: fix the method because it needs to stop throwing errors on every time
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
        public void sendKeysById(string elementId, string input) { //TODO: fix the method because it needs to stop throwing errors on every time
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
                        element.SendKeys(input);
                        print(method, " Keys sending" + "'"+input+"'");
                        break;
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
        public void addInputToEntryByClass(string elementName, string input) {
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
        public void clickByNameInTree(IWebElement parent, string elementName) {
            var element = findByName(parent, elementName);
            try {
                if (element.Enabled == true) {
                    element.Click();
                    debugLog.Info(method + " " + elementName + " Clicked");

                }
            } catch (NoSuchElementException e) {
                printError(method, null, e);

            }
        }
        public void clickByIdInTree(IWebElement parent, string elementName) {
            var element = findById(parent, elementName);
            try {
                if (element.Enabled == true) {
                    element.Click();
                    debugLog.Info(method + " " + elementName + " Clicked");
                }
            } catch (NoSuchElementException e) {
                printError(method, null, e);
            }
        }
        public IWebElement findById(IWebElement parent, string elementName) {
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
        public IWebElement findByName(IWebElement parent, string elementName) {
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

        public void closeDriver() {
            driver.Quit();
            print(method, "DRIVER CLOSED");
        }
        //Want this to cycle through all processes like kill method, 
        //but then cast to a window and get the windowState, then have it set the window max based off that 

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