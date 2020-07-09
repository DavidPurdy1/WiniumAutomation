using log4net;
using Microsoft.Office.Interop.Word;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Winium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace WiniumTests {
    public class UserMethods {
        Actions action;
        IWebElement window;
        WiniumDriver driver;
        string method;
        string mainWindowHandle;
        readonly DesktopOptions options = new DesktopOptions();
        readonly WiniumMethods m;
        readonly string driverPath;
        readonly ILog debugLog;

        public UserMethods(ILog log) {
            debugLog = log;
            options.ApplicationPath = ConfigurationManager.AppSettings.Get("IntactPath");
            driverPath = ConfigurationManager.AppSettings.Get("DriverPath");
            driver = new WiniumDriver(driverPath, options);
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            action = new Actions(driver);
            m = new WiniumMethods(driver, debugLog);
        }

        /**Doesn't work
         * This is going to connect to a remote server with the name brought in by serverName, when used
         * Doesn't work with the changes that have been made, lots of these are readonly
         */
        public void connectToRemoteDesktop(DesktopOptions options, string serverName) {

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
        public void loginToIntact() { //TODO: have to add connectToRemoteDesktop
            method = MethodBase.GetCurrentMethod().Name;
            debugLog.Info(method + " Started");
            //both of these will most likely stay false all the time, but we can test either of them by changing value in app.config
            bool needToSetDB = ConfigurationManager.AppSettings.Get("setDataBase") == "true"; ;
            bool connectToRemote = ConfigurationManager.AppSettings.Get("connectToRemote") == "true";
            Thread.Sleep(7000);
            m.SendKeys(By.Name(""), "admin");
            if (!needToSetDB) {
                m.Click(By.Name("&Logon"));
            } else {
                setDatabaseInformation();
                m.Click(By.Name("&Logon"));
            }
            Thread.Sleep(2000);
            mainWindowHandle = driver.CurrentWindowHandle;
            print(method, " Finished");
        }
        //UNTESTED
        private void setDatabaseInformation() {
            method = MethodBase.GetCurrentMethod().Name;
            debugLog.Info(method + " Started");
            m.Click(By.Name("&Settings.."));
            m.SendKeys(By.Name(""), @"(local)\INTACT");
            m.SendKeys(By.Name(""), "{TAB}");
            m.SendKeys(By.Name(""), "{TAB}");
            m.SendKeys(By.Name(""), "{ENTER}");

            debugLog.Info(method + " Finished");
        }
        /**This is going to a specified amount of definitions with random name for each blank.
         */
        public void createNewDefinition(int? numberOfDefinitions = 1) {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            window = driver.FindElement(By.Id("frmIntactMain"));
            window = window.FindElement(By.Name("radMenu1"));
            m.Click(By.Name("&Administration"), window);
            window = window.FindElement(By.Name("&Administration"));
            m.Click(By.Name("Definitions"), window);
            for (int i = 0; i <= numberOfDefinitions; i++) {
                var num = new Random().Next().ToString();
                window = m.Locate(By.Id("frmIntactMain"));
                window = m.Locate(By.Id("frmRulesList"));
                m.Click(By.Id("btnAdd"));
                window = m.Locate(By.Name("Add Definition"));
                print(method, "Definition name is " + "Test " + num);
                foreach (IWebElement element in window.FindElements(By.Name(""))) {
                    if (element.Enabled == true) {
                        try { element.SendKeys("Test " + num); } catch (Exception) { }
                    }
                }
                m.Click(By.Name("&Save"), window);
            }
            m.Click(By.Name("&Close"));
            print(method, "Finished");
        }
        /**Going to create a new type and will add random values for all of blanks.
         */
        public void createNewType(int? numberOfTypes = 1 ) {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");

            window = m.Locate(By.Id("frmIntactMain"));

            window = m.Locate(By.Name("radMenu1"));
            m.Click(By.Name("&Administration"), window);

            window = m.Locate(By.Name("&Administration"));
            m.Click(By.Name("Types"), window);

            for (int i = 0; i < numberOfTypes; i++) {
                var temp = new Random().Next().ToString();
                window = m.Locate(By.Id("frmIntactMain"));

                window = m.Locate(By.Id("frmAdminTypes"));

                m.Click(By.Id("rbtnAdd"), window);

                window = m.Locate(By.Id("frmAdminTypesInfo"));

                foreach (IWebElement element in window.FindElements(By.Name(""))) {
                    if (element.Enabled == true) {
                        try { element.SendKeys("Test " + temp); } catch (Exception) { }
                    }
                }
                m.Click(By.Name("&OK"));
            }
            m.Click(By.Name("&Close"));
            print(method, "Finished");
        }
        /**
        * Takes in isPDF telling whether or not it is a PDF or TIF and then numOfDocs for how many to add.
        * Going to take a random document from some preset path and with that it is going to add that document either pdf or not to a random definition and type with random metadata.
        */
        public void createDocument(int? numOfDocs = 1, bool isPDF = true, string docPath = "") {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");
            Thread.Sleep(2000);

            for (int i = 0; i < numOfDocs; i++) {
                m.Click(By.Name("Add Document"));

                Thread.Sleep(1000);
                //adding note 
                m.Click(By.Id("btnNotes"));
                m.Click(By.Id("rchkPrivate"));
                m.Click(By.Id("btnAddNote"));
                m.SendKeys(By.Id("txtNote"), "TEST NOTE");
                m.Click(By.Id("btnOK"));
                m.Click(By.Id("btnNotes"));

                //add document button (+ icon)
                print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
                m.Click(By.Id("lblType"));
                print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
                action.MoveByOffset(20, -40).Click().MoveByOffset(20, 60).Click().Build().Perform();
                print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);

                //find the document to add in file explorer
                //configure docpath in app.config, takes arg of pdf or tif 
                if (docPath.Length < 1) {
                    docPath = ConfigurationManager.AppSettings.Get("AddDocumentStorage");
                }
                m.SendKeys(By.Id("1001"), docPath);
                print(method, "Go to \"" + docPath + "\"");
                m.Click(By.Name("Go to \"" + docPath + "\""));

                var rand = new Random();
                if (isPDF) {
                    Winium.Elements.Desktop.ComboBox filesOfType = new Winium.Elements.Desktop.ComboBox(driver.FindElementByName("Files of type:"));
                    filesOfType.SendKeys("p");
                    action.KeyDown(OpenQA.Selenium.Keys.Alt).SendKeys("n").KeyUp(OpenQA.Selenium.Keys.Alt).SendKeys("PDF" + rand.Next(6).ToString()).Build().Perform();
                    m.Click(By.Id("1"));
                } else {
                    action.KeyDown(OpenQA.Selenium.Keys.Alt).SendKeys("n").KeyUp(OpenQA.Selenium.Keys.Alt).SendKeys("TIF" + rand.Next(6).ToString()).Build().Perform();
                    m.Click(By.Id("1"));
                }
                Thread.Sleep(3000);

                //rotate
                driver.FindElement(By.Id("lblType")).Click();
                action.MoveByOffset(375, -37).Click().Click().Build().Perform();

                //edit custom fields
                print(method, "custom fields");
                driver.FindElement(By.Id("lblType")).Click();
                action.MoveByOffset(150, 240).Click().SendKeys("1/1/2000").
                    MoveByOffset(0, 20).Click().SendKeys("7").MoveByOffset(0, 20).Click().SendKeys("string").Build().Perform();

                //edit fields
                print(method, "edit fields ");
                driver.FindElement(By.Id("lblType")).Click();
                action.MoveByOffset(170, 80).Click().SendKeys("1/1/2000").MoveByOffset(0, 20).Click().SendKeys("AUTHOR TEST").
                    MoveByOffset(0, 40).Click().SendKeys("SUMMARY TEST").Build().Perform();

                //save and quit
                print(method, "save and quit");
                m.Click(By.Id("btnSave"));
                m.Click(By.Id("btnClose"));
                print(method, "Finished");
            }
        }
        public bool search(string searchInput) {
            window = m.Locate(By.Id("frmIntactMain"));
            window = m.Locate(By.Name("radMenu1"), window);
            m.SendKeys(By.ClassName("WindowsForms10.EDIT.app.0.5c39d4"), searchInput, window);
            m.Click(By.Name("Search"), window);
            if (m.IsElementPresent(By.ClassName("WindowsForms10.MDICLIENT.app.0.5c39d4"))) {
                return true;
            }else if (m.IsElementPresent(By.Name("No search results found."))) {
                m.Click(By.Name("OK"));
                return false;
            } else {
                throw new Exception("Error in Search");
            }
        }
        public void openOrganizer() {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, " Started");
            m.Click(By.Name("Organizer"));
            print(method, " Finished");
        }
        /**
         * This method is going to add documents to batch review and then run through and both add a document to an existing document and attribute a new one.
         * Does usually run slow
         */
        public void BatchReview() {
            method = MethodBase.GetCurrentMethod().Name;
            addDocsToCollector();

            window = m.Locate(By.Id("frmIntactMain"));

            window = m.Locate(By.Id("radPanelBar1"), window);
            window = m.Locate(By.Id("pageIntact"), window);
            window = m.Locate(By.Id("lstIntact"), window);
            m.Click(By.Name("Batch Review"), window);
            m.Click(By.Id("6"));
            Thread.Sleep(1000);
            m.Click(By.Id("6"));


            //attribute test from batch review...
            BatchAttribution();

            //add to document test from batch review... 
            addDocBatchReview();

            m.Click(By.Id("btnClose"));
        }
        /**Should not call this method in tests
         */
        private void addDocBatchReview() {
            m.Click(By.Id("btnAddToDoc"));
            m.Click(By.Name("DEFAULT DEF"));
            m.Click(By.Name("DEFAULT DEFINITION TEST"));
            m.Click(By.Id("rbtnOK"));
            Thread.Sleep(2000);
            m.Locate(By.Name("&OK"));
            Thread.Sleep(2000);
            window = m.Locate(By.Id("frmInsertPagesVersion"));
            m.Click(By.Id("btnOK"), window);
            m.Locate(By.Id("frmDocument"));

            //rotate 
            print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            m.Click(By.Id("lblType"));
            print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            action.MoveByOffset(375, -37).Click().Click().Click().Click().Click().Build().Perform();
            print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);

            //custom fields
            print(method, "custom fields");
            driver.FindElement(By.Id("lblType")).Click();
            action.MoveByOffset(150, 240).Click().SendKeys("1/1/2000").
                MoveByOffset(0, 20).Click().SendKeys("10").MoveByOffset(0, 20).Click().SendKeys("BATCH REVIEW ADD").Build().Perform();

            //edit fields
            print(method, "edit fields ");
            driver.FindElement(By.Id("lblType")).Click();
            action.MoveByOffset(170, 80).Click().SendKeys("1/1/2000").MoveByOffset(0, 20).Click().SendKeys("BATCH AUTHOR TEST").
                MoveByOffset(0, 40).Click().SendKeys("BATCH ADDING TO ANOTHER DOCUMENT TEST").Build().Perform();

            m.Click(By.Name("Save"));
            m.Click(By.Name("Close"));
        }
        private void BatchAttribution() {
            window = m.Locate(By.Id("frmBatchReview"));
            m.Click(By.Id("btnAttribute"), window);
            Thread.Sleep(2000);
            window = m.Locate(By.Id("frmDocument"));

            //rotate 
            print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            m.Click(By.Id("lblType"));
            print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            action.MoveByOffset(375, -37).Click().Click().Build().Perform();
            print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);

            //custom fields
            print(method, "custom fields");
            driver.FindElement(By.Id("lblType")).Click();
            action.MoveByOffset(150, 240).Click().SendKeys("1/1/2000").
                MoveByOffset(0, 20).Click().SendKeys("10").MoveByOffset(0, 20).Click().SendKeys("BATCH REVIEW ATTRIBUTION").Build().Perform();
            
            //edit fields
            print(method, "edit fields ");
            driver.FindElement(By.Id("lblType")).Click();
            action.MoveByOffset(170, 80).Click().SendKeys("1/1/2000").MoveByOffset(0, 20).Click().SendKeys("BATCH AUTHOR TEST").
                MoveByOffset(0, 40).Click().SendKeys("BATCH ATTRIBUTING A DOCUMENT TEST").Build().Perform();

            m.Click(By.Id("btnSave"), window);
            m.Click(By.Id("btnClose"), window);
        }
        /**Collects documents by all definition from InZone. Returns true if the definition recognized is the same name as a file in directory
         */
        public bool getDocumentsFromInZone() {
            method = MethodBase.GetCurrentMethod().Name;
            addDocsToCollector();
            window = m.Locate(By.Id("frmIntactMain"));
            window = m.Locate(By.Id("radMenu1"), window);
            m.Click(By.Name("&Intact"), window);
            window = m.Locate(By.Name("&Intact"), window);
            m.Click(By.Name("InZone"), window);
            Thread.Sleep(2000);
            window = m.Locate(By.Id("frmInZoneMain"));
            m.Click(By.Id("btnCollectScan"), window);
            Thread.Sleep(9000);
            bool hasPassed = false;

            string startPath = ConfigurationManager.AppSettings.Get("InZoneStartPath");
            foreach (string s in Directory.GetFiles(startPath)) {
                string test = Path.GetFileName(s);
                if (m.IsElementPresent(By.Name(test.Substring(0, test.Length - 4)))) {
                    hasPassed = true;
                    break;
                }
            }
            m.Click(By.Id("btnCommit"), window);
            Thread.Sleep(1000);
            m.Click(By.Id("btnClose"), window);
            return hasPassed;
        }
        /**Should not call this method in tests
         * Method copies over files from one directory to another: Each time before InZone collects this is going to put files in the collector folder
         * Verify that the startPath always has files in it and those files shouldn't be removed from this folder when collected.
         */
        private void addDocsToCollector() {
            string startPath = ConfigurationManager.AppSettings.Get("InZoneStartPath");
            string endPath = ConfigurationManager.AppSettings.Get("InZoneCollectorPath");
            if (Directory.Exists(startPath) & Directory.Exists(endPath)) {
                foreach (string s in Directory.GetFiles(startPath)) {
                    File.Copy(s, Path.Combine(endPath, Path.GetFileName(s)), true);
                }
            } else {
                print(method, "Starting or Ending path doesn't exist");
            }
        }
        /** This is going to take a screenshot each time it is called and save it
         * 
         * TODO: MAKE IT DELETE THE IMAGES EACH TIME THAT WAY YOU DON'T GET MORE AND MORE IMAGES ON THE DOC FROM OTHER FAILED TESTS. TRY DOING THIS IN IMAGETODOC.
         */
        public string onFail(string testName) {
            string folderPath = ConfigurationManager.AppSettings.Get("AutomationScreenshots");
            string path = Path.Combine(folderPath, testName +"_" + DateTime.Now.Month.ToString() +"-" + DateTime.Now.Day.ToString() +"-" + DateTime.Now.Year.ToString() +
                "_" + DateTime.Now.Hour.ToString() +"-" + DateTime.Now.Minute.ToString()+"-" + DateTime.Now.Second.ToString());
            ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile(path, ImageFormat.Png);

            if (m.IsElementPresent(By.Id("frmDocument"))) {
                driver.Close();
                Thread.Sleep(2000);
            }
            else if (m.IsElementPresent(By.Id("frmInZoneMain"))) {
                driver.Close();
                Thread.Sleep(2000);
            }
            else if (m.IsElementPresent(By.Id("frmBatchReview"))) {
                driver.Close();
                Thread.Sleep(2000);
            }
            else if (m.IsElementPresent(By.Id("frmAdminTypesInfo"))) {
                driver.Close();
                Thread.Sleep(2000);
            }
            else if (m.IsElementPresent(By.Id("frmRulesEdit"))) {
                driver.Close();
                Thread.Sleep(2000);
            }
            else if (m.IsElementPresent(By.Id("frmZoneConfig"))) {
                driver.Close();
                Thread.Sleep(2000);
            }
            else if (m.IsElementPresent(By.Id("frmAdminTypesInfo"))) {
                driver.Close();
                Thread.Sleep(2000);
            }
            driver.SwitchTo().Window(mainWindowHandle);
            return path;
        }
        public void writeFailFile(List<string> testsFailedNames, List<string> testsPassedNames) {
            using (StreamWriter file =
            new StreamWriter(@"C:\Automation\failedTestsDocuments\FailedTests.txt", true)) {
                file.WriteLine(DateTime.Now.ToString() + "| " + testsFailedNames.Count.ToString() + " Tests failed | " + testsPassedNames.Count.ToString() + " Tests passed |");
                foreach (string name in testsFailedNames) {
                    file.WriteLine(name + " failed");
                }
                foreach (string name in testsPassedNames) {
                    file.WriteLine(name + " passed");
                }
                file.WriteLine(" ");
            }
        }
        /**Take images from folder and put them on a word doc
         * Put at the end of tests.
         */
        public void imageToDoc() { //have to figure out how to delete things afterwards
            method = MethodBase.GetCurrentMethod().Name;
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Document wordDoc = word.Documents.Add();
            Range range = wordDoc.Range();

            string folderPath = ConfigurationManager.AppSettings.Get("AutomationScreenshots");
            foreach (string image in Directory.GetFiles(folderPath)) {

                string imagePath = Path.Combine(folderPath, Path.GetFileName(image));
                print(method, imagePath);

                InlineShape autoScaledInlineShape = range.InlineShapes.AddPicture(imagePath);
                float scaledWidth = autoScaledInlineShape.Width;
                float scaledHeight = autoScaledInlineShape.Height;
                autoScaledInlineShape.Delete();

                // Create a new Shape and fill it with the picture
                Shape newShape = wordDoc.Shapes.AddShape(1, 0, 0, scaledWidth, scaledHeight);
                newShape.Fill.UserPicture(imagePath);

                // Convert the Shape to an InlineShape and optional disable Border
                InlineShape finalInlineShape = newShape.ConvertToInlineShape();
                finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                // Cut the range of the InlineShape to clipboard
                finalInlineShape.Range.Cut();

                // And paste it to the target Range
                range.Paste();
            }
            wordDoc.SaveAs2(ConfigurationManager.AppSettings.Get("TestFailedDocs"));
            word.Quit();
        }
        public void closeDriver() {
            m.closeDriver();
        }
        /** Have to add all scenarios where there is a new window open
         * This is to close out and get back to the window for the next test.
         */ 

        /**For debug
         */
        private void print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
        private void printError(string method, string toPrint, Exception e) {
            debugLog.Info(method + " " + e + " " + toPrint);
        }
        private void printCursor() {
            print(method, "x: " + Cursor.Position.X.ToString() + " y: " + Cursor.Position.Y.ToString());
        }
    }
}