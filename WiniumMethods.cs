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
            } catch (NoSuchElementException e) {
                print(method, " Failed on " + method + "Finding element" + by.ToString() + e.Source);
                throw new AssertFailedException("Failed on " + method + "Finding element" + by.ToString());
            }
            print(method, "Was not NoSuchElement");
            throw new AssertFailedException(method + " Was not NoSuchElement");
        }
        public IWebElement Locate(By by, IWebElement parent) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                IWebElement element = parent.FindElement(by);
                print(method, by.ToString() + " Has been Located");
                if (element != null) {
                    return element;
                }
            } catch (NoSuchElementException e) {
                print(method, " Failed on " + method + " Finding element" + by.ToString() + e.Source);
                throw new AssertFailedException("Failed on " + method + " Finding element" + by.ToString());
            }
            print(method, "Was not NoSuchElement");
            throw new AssertFailedException(method + " Was not NoSuchElement");
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
            } catch (ElementNotVisibleException) {
                print(method, "Could not click element: Not on Screen " + by.ToString());
                throw new AssertFailedException(method + "Could not click element: Not on Screen " + by.ToString());
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
            } catch (ElementNotVisibleException) {
                print(method, "Could not click element: Not on Screen " + by.ToString());
                throw new AssertFailedException(method + "Could not click element: Not on Screen " + by.ToString());
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
            } catch (ElementNotVisibleException) {
                print(method, "Could not send to element: Not on Screen " + by.ToString());
                throw new AssertFailedException(method + "Could not send element: Not on Screen " + by.ToString());
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
            } catch (ElementNotVisibleException) {
                print(method, "Could not send to element: Not on Screen " + by.ToString());
                throw new AssertFailedException(method + "Could not send element: Not on Screen " + by.ToString());
            }
        }
        public bool IsElementPresent(By by) {
            return driver.FindElements(by).Count > 0;
        }
        public void closeDriver() {
            driver.Quit();
            print(method, "DRIVER CLOSED");
        }

        private void print(string method, string toPrint = "", Exception e = null) {
            debugLog.Info(method + " " + toPrint + " " + e);
        }
    }
}