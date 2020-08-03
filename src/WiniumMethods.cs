using log4net;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Reflection;


namespace WiniumTests.src {
    /// <summary>
    /// Class containing methods used to locate elements, interact with the form, and driver actions
    /// </summary>
    public class WiniumMethods {
        string method;
        public readonly WindowsDriver<WindowsElement> driver;
        readonly ILog debugLog;
        public string mainHandle;
        public List<string> handleList = new List<string>();
        public WiniumMethods(WindowsDriver<WindowsElement> driver, ILog log) {
            this.driver = driver;
            debugLog = log;
        }

        public IWebElement Locate(By by) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                Print(method, "test");
                if (GetFindMethod(by) == "By.Id") {
                    Print(method, GetElement(by));
                    IWebElement element = driver.FindElementByAccessibilityId(GetElement(by));
                    Print(method, by.ToString() + " Has been Located");
                    return element;
                } else {
                    IWebElement element = driver.FindElement(by);
                    Print(method, by.ToString() + " Has been Located");
                    return element;
                }
            } catch (NoSuchElementException e) {
                Print(method, " Failed on " + method + " Finding element" + by.ToString() + e.StackTrace);
                throw new AssertFailedException("Failed on " + method + " Finding element" + by.ToString());
            }
        }
        public IWebElement Locate(By by, IWebElement parent) {
            method = MethodBase.GetCurrentMethod().Name;
            try {
                if (GetFindMethod(by) == "By.Id") {
                    var v = (WindowsElement)parent; 
                    var element = v.FindElementByAccessibilityId(GetElement(by));
                    Print(method, by.ToString() + " Has been Located");
                    return element;
                } else {
                    IWebElement element = parent.FindElement(by);
                    Print(method, by.ToString() + " Has been Located");
                    return element;
                }
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
            } catch (StaleElementReferenceException e) {
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
            } catch (StaleElementReferenceException e) {
                Print(method, "Could not send to element " + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + "Could not send to element " + by.ToString());
            } catch (InvalidElementStateException e) {
                Print(method, " element is not able to recieve keys" + by.ToString() + e.StackTrace);
                throw new AssertFailedException(method + " element is not able to recieve keys" + by.ToString());
            } 
            //catch (ElementNotVisibleException e) {
            //    Print(method, "Could not send to element: Not on Screen " + by.ToString() + e.StackTrace);
            //    throw new AssertFailedException(method + "Could not send element: Not on Screen " + by.ToString());
            //}
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
            } 
            //catch (ElementNotVisibleException e) {
            //    Print(method, "Could not send to element: Not on Screen " + by.ToString() + e.StackTrace);
            //    throw new AssertFailedException(method + "Could not send element: Not on Screen " + by.ToString());
            //}
        }
        public bool IsElementPresent(By by) {
            return driver.FindElements(by).Count > 0;
        }
        public bool IsElementPresent(By by, IWebElement parent) {
            return parent.FindElements(by).Count > 0;
        }
        public void SwitchWindowHandle(int i = 0) {
            try {
                driver.SwitchTo().Window(driver.WindowHandles[i]);
            } catch (NoSuchWindowException e ) {
                Console.WriteLine("Failed on Switching the window " + e.StackTrace);
            }
        }
        public Screenshot GetScreenshot() {
            return ((ITakesScreenshot)driver).GetScreenshot();
        }
        public void CloseDriver() {
            driver.Quit();
            Print(method, "DRIVER CLOSED");
        }
        private string GetFindMethod(By by) {
            var locatorMethod = by.ToString();
            if (!string.IsNullOrWhiteSpace(locatorMethod)) {
                int semiLocation = locatorMethod.IndexOf(":",StringComparison.Ordinal);
                if (semiLocation > 0) { 
                    return locatorMethod.Substring(0, semiLocation); 
                }
            }
            return string.Empty;
        }
        private string GetElement(By by) {
            var locatorMethod = by.ToString();
            if (!string.IsNullOrWhiteSpace(locatorMethod)) {
                int semiLocation = locatorMethod.IndexOf(":", StringComparison.Ordinal);
                if (semiLocation > 0) {
                    return locatorMethod.Substring(semiLocation + 2 , locatorMethod.Length-semiLocation-2);
                }
            }
            return string.Empty;
        }
        public string GetCurrentHandle() {
            return driver.CurrentWindowHandle; 
        }
        public void LoopWindowHandle() {
            int i = 0;
            while(mainHandle == driver.CurrentWindowHandle) {
                SwitchWindowHandle(i);
                i++;
            }
        }
        public void NextHandle() {
            foreach(var handle in driver.WindowHandles) {
                foreach(var stored in handleList) {
                    if(!handle.Equals(stored)) {
                        driver.SwitchTo().Window(handle);
                        return; 
                    }
                }
            }
        }
        public void WindowCount() {
            for (int i = 0; i < driver.WindowHandles.Count; i++) {
                Print(method, i.ToString());
            }
        }
        public void For() {
            for (int i = 0; i < driver.WindowHandles.Count; i++) {
                string handle = driver.WindowHandles[driver.WindowHandles.Count - i];
                driver.SwitchTo().Window(handle);
            }
        }
        private void Print(string method, string toPrint = "", Exception e = null) {
            debugLog.Info(method + " " + toPrint + " " + e);
        }
    }
}