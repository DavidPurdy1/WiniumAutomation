using log4net;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Winium;
using System;
using System.Reflection;


namespace WiniumTests {
    public class WiniumMethods {
        string method;
        readonly WiniumDriver driver;
        readonly ILog debugLog;

        public WiniumMethods(WiniumDriver driver, ILog log) {
            this.driver = driver;
            debugLog = log;
        }


        public IWebElement Locate(By by) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                IWebElement element = driver.FindElement(by);
                print(method, by.ToString() + " Has been Located");
                if (element != null) {
                    return element;
                }
                throw new ElementNullException();
            } catch (NoSuchElementException e) {
                print(method, " Failed on " + method + "Finding element" + by.ToString() + e.Source);
                throw new AssertFailedException("Failed on " + method + "Finding element" + by.ToString());
            }
            throw new ElementNullException();
        }
        public IWebElement Locate(By by, IWebElement parent) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                IWebElement element = parent.FindElement(by);
                print(method, by.ToString() + " Has been Located");
                if (element != null) {
                    return element;
                }
                throw new ElementNullException();
            } catch (NoSuchElementException e) {
                print(method, " Failed on " + method + " Finding element" + by.ToString() + e.Source);
                throw new AssertFailedException("Failed on " + method + " Finding element" + by.ToString());
            }
            throw new ElementNullException();
        }
        public void Click(By by) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                var element = Locate(by);
                if (element.Enabled) {
                    element.Click();
                    print(method, by.ToString() + " Clicked");
                }
            } catch (StaleElementReferenceException) {
                print(method, "Could not click element " + by.ToString());
                throw new AssertFailedException(method + "Could not click element " + by.ToString());
            }
        }
        public void Click(By by, IWebElement parent) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                var element = Locate(by, parent);
                if (element.Enabled) {
                    element.Click();
                    print(method, by.ToString() + " Clicked");
                }
            } catch (StaleElementReferenceException) {
                print(method, "Could not click element " + by.ToString());
                throw new AssertFailedException(method + "Could not click element " + by.ToString());
            }
        }
        public void SendKeys(By by, string input) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                var element = Locate(by);
                if (element.Enabled) {
                    element.SendKeys(input);
                    print(method, by.ToString() + " sent " + input);
                }
            } catch (StaleElementReferenceException) {
                print(method, "Could not send to element " + by.ToString());
                throw new AssertFailedException(method + "Could not send to element " + by.ToString());
            } catch (InvalidElementStateException) {
                print(method, " element is not able to recieve keys" + by.ToString());
                throw new AssertFailedException(method + " element is not able to recieve keys" + by.ToString());
            }
        }
        public void SendKeys(By by, string input, IWebElement parent) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                var element = Locate(by, parent);
                if (element.Enabled) {
                    element.SendKeys(input);
                    print(method, by.ToString() + " sent " + input);
                }
            } catch (StaleElementReferenceException) {
                print(method, "Could not send to element " + by.ToString());
                throw new AssertFailedException(method + "Could not send to element " + by.ToString());
            } catch (InvalidElementStateException) {
                print(method, " element is not able to recieve keys" + by.ToString());
                throw new AssertFailedException(method + " element is not able to recieve keys" + by.ToString());
            }
        }
        public bool IsElementPresent(By by) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                Locate(by);
                print(method, " Element is present");
                return true;
            } catch (NoSuchElementException) {
                return false;
            }
        }
        public void closeDriver() {
            driver.Quit();
            print(method, "DRIVER CLOSED");
        }


        public void print(string method, string toPrint = "", Exception e = null) {
            debugLog.Info(method + " " + toPrint + " " + e);
        }
        public void PrintError(string method, Exception e, string toPrint = "") {
            debugLog.Info(method + " " + e + " " + toPrint);
        }

    }
}