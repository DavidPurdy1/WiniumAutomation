using log4net;
using Microsoft.Office.Interop.Word;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Winium;
using System;
using System.Configuration;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Winium.Elements.Desktop.Extensions;

namespace WiniumTests {
    public class UserMethods { //TODO: THE INSTANCE OF THE DRIVER IS WERIRD FIX IT!! Has to be logged in at least once for it to work
        //Add Actions globally so it can be used over again, maybe window as long as it is all good
        WiniumDriver driver;
        DesktopOptions options = new DesktopOptions();
        WiniumMethods m;
        bool needToSetDB = false;
        string driverPath;
        string method;
        ILog debugLog;

        public UserMethods(ILog log) {
            debugLog = log;
            options.ApplicationPath = ConfigurationManager.AppSettings.Get("IntactPath");
            driverPath = ConfigurationManager.AppSettings.Get("DriverPath");
            driver = new WiniumDriver(driverPath, options);
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            m = new WiniumMethods(driver, debugLog);
        }

        /**
         * This is going to connect to a remote server with the name brought in by serverName, when used
         * Doesn't work with the changes that have been made
         */
        public void connectToRemoteDesktop(DesktopOptions options, string serverName) {

            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");

            options.ApplicationPath = ConfigurationManager.AppSettings.Get("RemoteDesktop"); 
            driver = new WiniumDriver(driverPath, options);
            m.sendKeysByName("Remote Desktop Connection", serverName);
            m.clickByName("Connect");

            print(method, "Finished");
        }
        /**
         * Logs into intact automatically with the admin login, THIS METHOD HAS TO BE RAN FIRST WITH THE DRIVER 
         */
        public void loginToIntact() {
            method = MethodBase.GetCurrentMethod().Name;
            debugLog.Info(method + " Started");
            Thread.Sleep(7000);
            m.sendKeysById("", "admin");
            if (!needToSetDB) {
                m.clickByName("&Logon");
            } else {
                setDatabaseInformation();
                m.clickByName("&Logon");
            }
            debugLog.Info(method + " Finished");
        }
        //UNTESTED
        public void setDatabaseInformation() {
            method = MethodBase.GetCurrentMethod().Name;
            debugLog.Info(method + " Started");
            m.clickByName("&Settings..");
            m.sendKeysByName("", @"(local)\INTACT");
            m.sendKeysByName("", "{TAB}");
            m.sendKeysByName("", "{TAB}");
            m.sendKeysByName("", "{ENTER}");

            debugLog.Info(method + " Finished");
        }
        /**
         * This is going to a specified amount of definitions with random name for each blank. TODO: RIGHT NOW IT IS IMPORTANT THAT INTACT IS FULLSCREEN WILL NOT WORK OTHERWISE
         */
        public void createNewDefinition(int? numberOfDefinitions = 1) {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");

            var window = driver.FindElement(By.Id("frmIntactMain"));

            window = window.FindElement(By.Name("radMenu1"));
            m.clickByNameInTree(window, "&Administration");

            window = window.FindElement(By.Name("&Administration"));
            m.clickByNameInTree(window, "Definitions");

            for (int i = 0; i <= numberOfDefinitions; i++) {
                var num = new Random().Next().ToString();

                window = driver.FindElementById("frmIntactMain");

                window = window.FindElement(By.Id("frmRulesList"));

                window.FindElement(By.Id("btnAdd")).Click();

                window = window.FindElement(By.Name("Add Definition"));

                debugLog.Info("Definition name is " + "Test " + num);

                foreach (IWebElement element in window.FindElements(By.Name(""))) {
                    if (element.Enabled == true) {
                        try { element.SendKeys("Test " + num); } catch (Exception) { }
                    }
                }
                m.clickByNameInTree(window, "&Save");
            }
            m.clickByName("&Close");
            print(method, "Finished");
        }
        /**
         * Going to create a new type and will add random values for all of blanks. VERY IMPORTANT THAT THE INTACT WINDOW IS FULLSCREENED FOR RIGHT NOW 
         */
        public void createNewType(int? numberOfTypes = 1 ) {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");

            var window = driver.FindElement(By.Id("frmIntactMain"));

            window = window.FindElement(By.Name("radMenu1"));

            m.clickByNameInTree(window, "&Administration");

            window = window.FindElement(By.Name("&Administration")); //add the setting of the window in the tree method and rename it
            m.clickByNameInTree(window, "Types");


            for (int i = 0; i < numberOfTypes; i++) {
                var temp = new Random().Next().ToString();
                window = driver.FindElement(By.Id("frmIntactMain"));

                window = window.FindElement(By.Id("frmAdminTypes"));
                window.FindElement(By.Id("rbtnAdd")).Click();

                window = window.FindElement(By.Id("frmAdminTypesInfo"));

                foreach (IWebElement element in window.FindElements(By.Name(""))) {
                    if (element.Enabled == true) {
                        try { element.SendKeys("Test " + temp); } catch (Exception) { }
                    }
                }
                m.clickByName("&OK");
            }
            m.clickByName("&Close");
            print(method, "Finished");
        }
        /** //TODO: Need to refine createDocument a bit
        * Takes in isPDF telling whether or not it is a PDF or TIF and then numOfDocs for how many to add.
        * Going to take a random document from some preset path and with that it is going to add that document either pdf or not to a random definition and type with random metadata.
        */
        public void createDocument(int? numOfDocs = 1, bool isPDF = true, string docPath = @"C:\Automation\createDocumentStorage") {
            method = MethodBase.GetCurrentMethod().Name;
            print(method, "Started");

            driver.FindElement(By.Name("Add Document")).Click();
            driver.FindElementById("btnNotes").Click();
            driver.FindElementById("btnAddNote").Click();
            driver.FindElementById("txtNote").SendKeys("TEST NOTE");
            driver.FindElementById("btnOK").Click();
            driver.FindElementById("btnNotes").Click();


            print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            driver.FindElement(By.Id("lblType")).Click();
            print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            Actions action = new Actions(driver);
            action.MoveByOffset(20, -40).Click().MoveByOffset(20,60).Click().Build().Perform();
            print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);

            //m.sendKeysById("radCommandBar1", Keys.Control + "a");    //alternative method if coordinates don't work 

            m.sendKeysById("1001", docPath);
            print(method, "Go to \"" + docPath + "\"");
            driver.FindElement(By.Name("Go to \""+ docPath + "\"")).Click();

            for (int i = 0; i < numOfDocs; i++) {
                var rand = new Random();
                if (isPDF) {
                    Winium.Elements.Desktop.ComboBox filesOfType = new Winium.Elements.Desktop.ComboBox(driver.FindElementByName("Files of type:"));
                    filesOfType.SendKeys("p");
                    action.KeyDown(OpenQA.Selenium.Keys.Alt).SendKeys("n").KeyUp(OpenQA.Selenium.Keys.Alt).SendKeys("PDF" + rand.Next(3).ToString()).Build().Perform();
                    driver.FindElementById("1").Click();
                } else {
                    action.KeyDown(OpenQA.Selenium.Keys.Alt).SendKeys("n").KeyUp(OpenQA.Selenium.Keys.Alt).SendKeys("TIF" + rand.Next(3).ToString()).Build().Perform();
                }
            }
            Thread.Sleep(3000);
            driver.FindElement(By.Id("btnSave")).Click();
            driver.FindElement(By.Id("btnClose")).Click();

            //var window = driver.FindElementById("pnlAttributes");
            //m.findById(window, "dtcExpireDate").SendKeys("1/1/3000");
            //m.findById(window, "txtAuthor").SendKeys("AUTHOR TEST");
            //m.findById(window, "txtSummary").SendKeys("SUMMARY TEST");

            print(method, "Finished");
        }
        public void openOrganizer() {
            method = MethodBase.GetCurrentMethod().Name;
            debugLog.Info(method + " Started");
            m.clickByName("Organizer");
            debugLog.Info(method + " Finished");
        }
        /**
         * This method is going to add documents to batch review and then run through and both add a document to an existing document and attribute a new one.
         */
        public void BatchReview() {
            addDocsToCollector();
            var window = driver.FindElement(By.Id("frmIntactMain"));
            window = m.findById(window, "radPanelBar1");
            window = m.findById(window, "pageIntact");
            window = m.findById(window, "lstIntact");
            m.clickByNameInTree(window, "Batch Review");
            driver.FindElement(By.Id("6")).Click();
            driver.FindElement(By.Id("6")).Click();

            //window = driver.FindElement(By.Name("Review"));
            //window = m.findById(window, "radSplitContainer1");
            //window = driver.FindElement(By.Id("splitPanel2"));
            // window = m.findById(window, "radCommandBar1");
            //m.clickByIdInTree(window, "btnRotate");
            //m.getDriver().FindElementById("btnRotate").Click();
            //string[] zoomInput = { "Fit Width", "Fit Height", "10%", "50%", "100%", "200%", "400%" };
            // for (int t = 0; t < zoomInput.Length; t++) {
            //  m.sendKeysById("329982", zoomInput[t] + Keys.Enter);
            //}
            // print(method, window.Location.X.ToString() + " " + window.Location.Y.ToString());
            //driver.FindElementById("btnRotate").Click();

            //var element = m.findById(window, "radCommandBar1");
            //SelectElement select = new SelectElement(element);
            //select.SelectByIndex(4);

            //attribute test from batch review...
            driver.FindElementById("btnAttribute").Click();
            driver.FindElementById("btnSave").Click();
            driver.FindElementById("btnClose").Click();

            //add to document test from batch review... 
            addDocBatchReview();

            driver.FindElementById("btnClose").Click();


        }
        private void addDocBatchReview() {
            driver.FindElementById("btnAddToDoc").Click();
            driver.FindElementByName("DEFAULT DEF").Click();
            driver.FindElementByName("DEFAULT DEFINITION TEST").Click();
            driver.FindElementById("rbtnOK").Click();
            driver.FindElement(By.Name("&OK"));
            var window = driver.FindElementById("frmInsertPagesVersion");
            m.clickByIdInTree(window, "btnOK");
            driver.FindElementByName("Save").Click();
            driver.FindElementByName("Close").Click();
        }
        /**
         * Collects documents by all definition from InZone.
         * 
         * 
         * 
         * TODO: Add if inzone doesn't recognize the definition come in fail the test
         */
        public bool getDocumentsFromInZone() { //This is supposed to return true if the document has been identified, UNTESTED RIGHT NOW
            addDocsToCollector();
            var window = driver.FindElement(By.Id("frmIntactMain"));
            window = m.findById(window, "radMenu1");
            m.clickByNameInTree(window, "&Intact");
            window = m.findByName(window, "&Intact");
            m.clickByNameInTree(window, "InZone");
            window = driver.FindElement(By.Id("frmInZoneMain"));
            m.clickByIdInTree(window, "btnCollectScan");
            Thread.Sleep(5000);
            bool hasPassed = false;
            print(method, "THIS IS HOW MANY ROWS" + m.findById(window, "grdDocs").ToDataGrid().RowCount.ToString());
            if (m.findByName(window, "CleanFreak InZone Test").Displayed) {
                hasPassed = true;
            }
            m.clickByIdInTree(window, "btnCommit");
            m.clickByIdInTree(window, "btnClose");
            return hasPassed;
        }
        /**
         * Method copies over files from one directory to another: Each time before InZone collects this is going to put files in the collector folder
         * Verify that the startPath always has files in it and those files shouldn't be removed from this folder when collected
         * 
         * TODO: ADD AN EVENT TIMER WHERE ONCE THE WINDOW IN FOCUS DISPOSES YOU COULD PROBABLY NOW CLICK COMMIT DOCUMENTS
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
        /**
         * This is going to take a screenshot each time it is called and save it to the folder where the imageToDoc() method will save it and put it on a word doc.
         * 
         * TODO: MAKE IT DELETE THE IMAGES EACH TIME THAT WAY YOU DON'T GET MORE AND MORE IMAGES ON THE DOC FROM OTHER FAILED TESTS. TRY DOING THIS IN IMAGETODOC.
         */
        public void failLog() {
            string folderPath = ConfigurationManager.AppSettings.Get("AutomationScreenshots");
            var rand = new Random();
            var path = Path.Combine(folderPath, "test " + rand.Next(100).ToString());
            ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile(path, ImageFormat.Png);

        }
        /**
         * Going to grab the images in the folder that has failed and is going to put them on a word doc. Put at the end of tests.
         */
        public void imageToDoc() { //have to figure out how to delete things afterwards
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