using log4net;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Winium;
using System;
using System.Reflection;


namespace WiniumTests.src {
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
                Print(method, by.ToString() + " Has been Located");
                return element;
            } catch (NoSuchElementException e) {
                Print(method, " Failed on " + method + " Finding element" + by.ToString() + e.StackTrace);
                throw new AssertFailedException("Failed on " + method + " Finding element" + by.ToString());
            }
        }
        public IWebElement Locate(By by, IWebElement parent) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                IWebElement element = parent.FindElement(by);
                Print(method, by.ToString() + " Has been Located");
                return element;
            } catch (NoSuchElementException e) {
                Print(method, " Failed on " + method + " Finding element" + by.ToString() + e.StackTrace);
                throw new AssertFailedException("Failed on " + method + " Finding element" + by.ToString());
            }
        }
        public void Click(By by) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                var element = Locate(by);
                if (element.Enabled) {
                    element.Click();
                    Print(method, by.ToString() + " Clicked");
                }
            } catch (StaleElementReferenceException e ) {
                Print(method, "Could not click element " + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + "Could not click element " + by.ToString());
            } catch (ElementNotVisibleException e) {
                Print(method, "Could not click element: Not on Screen " + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + "Could not click element: Not on Screen " + by.ToString());
            }
        }
        public void Click(By by, IWebElement parent) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                var element = Locate(by, parent);
                if (element.Enabled) {
                    element.Click();
                    Print(method, by.ToString() + " Clicked");
                }
            } catch (StaleElementReferenceException e) {
                Print(method, "Could not click element " + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + "Could not click element " + by.ToString());
            } catch (ElementNotVisibleException e) {
                Print(method, "Could not click element: Not on Screen " + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + "Could not click element: Not on Screen " + by.ToString());
            }
        }
        public void SendKeys(By by, string input) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                var element = Locate(by);
                if (element.Enabled) {
                    element.SendKeys(input);
                    Print(method, by.ToString() + " sent " + input);
                }
            } catch (StaleElementReferenceException e ) {
                Print(method, "Could not send to element " + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + "Could not send to element " + by.ToString());
            } catch (InvalidElementStateException e ) {
                Print(method, " element is not able to recieve keys" + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + " element is not able to recieve keys" + by.ToString());
            } catch (ElementNotVisibleException e) {
                Print(method, "Could not send to element: Not on Screen " + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + "Could not send element: Not on Screen " + by.ToString());
            }
        }
        public void SendKeys(By by, string input, IWebElement parent) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                var element = Locate(by, parent);
                if (element.Enabled) {
                    element.SendKeys(input);
                    Print(method, by.ToString() + " sent " + input);
                }
            } catch (StaleElementReferenceException e) {
                Print(method, "Could not send to element " + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + "Could not send to element " + by.ToString());
            } catch (InvalidElementStateException e) {
                Print(method, " element is not able to recieve keys" + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + " element is not able to recieve keys" + by.ToString());
            } catch (ElementNotVisibleException e) {
                Print(method, "Could not send to element: Not on Screen " + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + "Could not send element: Not on Screen " + by.ToString());
            }
        }
        public bool IsElementPresent(By by) {
            return driver.FindElements(by).Count > 0;
        }
        public bool IsElementPresent(By by, IWebElement parent) {
            return parent.FindElements(by).Count > 0;
        }
        private void Print(string method, string toPrint = "", Exception e = null) {
            debugLog.Info(method + " " + toPrint + " " + e);
        }
    }
}