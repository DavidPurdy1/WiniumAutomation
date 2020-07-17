using log4net;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
        #region 
        IWebElement window;
        string method;
        readonly DesktopOptions options = new DesktopOptions();
        readonly WiniumDriver driver;
        readonly WiniumMethods m;
        readonly ILog debugLog;
        readonly Actions action;
        #endregion

        #region
        public UserMethods(ILog log) {
            debugLog = log;
            options.ApplicationPath = ConfigurationManager.AppSettings.Get("IntactPath");
            driver = new WiniumDriver(ConfigurationManager.AppSettings.Get("DriverPath"), options);
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            action = new Actions(driver);
            m = new WiniumMethods(driver, debugLog);
        }
        /**Doesn't work
         * This is going to connect to a remote server with the name brought in by serverName, when used
         * Doesn't work with the changes that have been made, lots of these are readonly
         */
        public void ConnectToRemoteDesktop(DesktopOptions options, string serverName) {

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
        public void LoginToIntact() { //TODO: have to add connectToRemoteDesktop
            method = MethodBase.GetCurrentMethod().Name;
            debugLog.Info(method + " Started");
            //both of these will most likely stay false all the time, but we can test either of them by changing value in app.config
            bool needToSetDB = ConfigurationManager.AppSettings.Get("setDataBase") == "true"; ;
            bool connectToRemote = ConfigurationManager.AppSettings.Get("connectToRemote") == "true";
            Thread.Sleep(10000);
            m.SendKeys(By.Name(""), "admin", m.Locate(By.Id("frmLogin")));
            if (!needToSetDB) {
                m.Click(By.Name("&Logon"));
            } else {
                setDatabaseInformation();
                m.Click(By.Name("&Logon"));
            }
            Thread.Sleep(2000);
            Print(method, " Finished");
        }
        public void Logout() {
            window = m.Locate(By.Name("&Intact"), m.Locate(By.Name("radMenu1")));
            m.Click(By.Name("Log Out"), window);
        }
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
        #endregion

        #region
        //TO TEST THIS MAKE SURE THAT INTACT IS IN FULL SCREEN, FIX

        /**This is going to a specified amount of definitions with random name for each blank.
         */
        public void CreateNewDefinition(int? numberOfDefinitions = 1, string definitionName = "") {
            method = MethodBase.GetCurrentMethod().Name;
            Print(method, "Started");

            window = m.Locate(By.Id("frmIntactMain"));
            window = m.Locate(By.Name("radMenu1"),window);
            m.Click(By.Name("&Administration"), window);
            window = m.Locate(By.Name("&Administration"),window);
            m.Click(By.Name("Definitions"), window);

            if(definitionName.Length < 2) {
                definitionName = "Test";
            }

            for (int i = 0; i <= numberOfDefinitions; i++) {
                var num = new Random().Next().ToString();
                window = m.Locate(By.Id("frmRulesList"), m.Locate(By.Id("frmIntactMain")));
                m.Click(By.Id("btnAdd"), window);
                window = m.Locate(By.Name("Add Definition"));
                Print(method, "Definition name is " + definitionName + num);
                foreach (IWebElement element in window.FindElements(By.Name(""))) {
                    if (element.Enabled == true) {
                        try { element.SendKeys(definitionName+ " " + num); } catch (Exception) { }
                    }
                }
                m.Click(By.Name("&Save"), window);
            }

            m.Click(By.Name("&Close"));
            Print(method, "Finished");
        }
        /**Going to create a new type and will add random values for all of blanks.
         * numberOfTypes: specify how many you want to add
         */
        public void CreateNewType(int? numberOfTypes = 1, string typeName = "") {
            method = MethodBase.GetCurrentMethod().Name;
            Print(method, "Started");

            window = m.Locate(By.Name("radMenu1"), m.Locate(By.Id("frmIntactMain")));
            m.Click(By.Name("&Administration"), window);

            m.Click(By.Name("Types"), m.Locate(By.Name("&Administration")));

            if(typeName.Length < 2) {
                typeName = "Test";
            }
            for (int i = 0; i < numberOfTypes; i++) {
                var temp = new Random().Next().ToString();

                window = m.Locate(By.Id("frmAdminTypes"), m.Locate(By.Id("frmIntactMain")));

                m.Click(By.Id("rbtnAdd"), window);
                window = m.Locate(By.Id("frmAdminTypesInfo"));

                foreach (IWebElement element in window.FindElements(By.Name(""))) {
                    if (element.Enabled == true) {
                        try { element.SendKeys(typeName + temp); } catch (Exception) { }
                    }
                }
                m.Click(By.Name("&OK"));
            }
            m.Click(By.Name("&Close"));
            Print(method, "Finished");
        }
        /**Creates docs
        * numOfDocs: specifies how many to create
        * if isPDF = true --> gets pdf from the directory, else tif 
        * docPath: allows you to specify the directory of docs, default is set in config
        * fileNumber: allows you to specify which file you want to use
        */
        public void CreateDocument(int? numOfDocs = 1, bool isPDF = true, string docPath = "", int? fileNumber = 0) {
            method = MethodBase.GetCurrentMethod().Name;
            Print(method, "Started");
            Thread.Sleep(2000);

            for (int i = 0; i < numOfDocs; i++) {
                m.Click(By.Name("Add Document"));

                Thread.Sleep(1000);
                //adding note 
                m.Click(By.Id("btnNotes"));
                m.Click(By.Id("btnAddNote"));
                m.Click(By.Id("rchkPrivate"));
                m.SendKeys(By.Id("txtNote"), "TEST NOTE");
                m.Click(By.Id("btnOK"));
                m.Click(By.Id("btnNotes"));

                //add document button (+ icon)
                Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
                m.Click(By.Id("lblType"));
                Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
                action.MoveByOffset(20, -40).Click().MoveByOffset(20, 60).Click().Build().Perform();
                Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);

                //find the document to add in file explorer
                //configure docpath in app.config, takes arg of pdf or tif 
                if (docPath.Length < 1) {
                    docPath = ConfigurationManager.AppSettings.Get("AddDocumentStorage");
                }
                m.SendKeys(By.Id("1001"), docPath);
                Print(method, "Go to \"" + docPath + "\"");
                m.Click(By.Name("Go to \"" + docPath + "\""));

                var rand = new Random();
                if (isPDF) {
                    Winium.Elements.Desktop.ComboBox filesOfType = new Winium.Elements.Desktop.ComboBox(m.Locate(By.Name("Files of type:")));
                    filesOfType.SendKeys("p");
                    filesOfType.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(500);
                    if (fileNumber == 0) {
                        action.MoveToElement(m.Locate(By.Id(rand.Next(Directory.GetFiles(docPath, "*.pdf").Length).ToString()))).DoubleClick().Build().Perform();
                    } else {
                        action.MoveToElement(m.Locate(By.Id(fileNumber.ToString()))).DoubleClick().Build().Perform();
                    }
                    m.Click(By.Name("Open"));
                } else {
                    if (fileNumber == 0) {
                        action.MoveToElement(m.Locate(By.Id(rand.Next(Directory.GetFiles(docPath, "*.tif").Length).ToString()))).DoubleClick().Build().Perform();
                    } else {
                        action.MoveToElement(m.Locate(By.Id(fileNumber.ToString()))).DoubleClick().Build().Perform();
                    }
                    m.Click(By.Name("Open"));
                }
                Thread.Sleep(3000);

                //rotate
                //m.Click(By.Id("lblType"));
                //action.MoveByOffset(375, -37).Click().Click().Build().Perform();

                //add annotations: UNABLE TO DO THIS WITHOUT DRAG AND DROP. NOT IMPLEMENTED EXCEPTION
                //AddAnnotations();

                //edit custom fields
                Print(method, "custom fields");
                m.Click(By.Id("lblType"));
                action.MoveByOffset(150, 240).Click().SendKeys("1/1/2000").
                    MoveByOffset(0, 20).Click().SendKeys("7").MoveByOffset(0, 20).Click().SendKeys("string").Build().Perform();

                //edit fields
                Print(method, "edit fields ");
                m.Click(By.Id("lblType"));
                action.MoveByOffset(170, 80).Click().SendKeys("1/1/2000").MoveByOffset(0, 20).Click().SendKeys("AUTHOR TEST").
                    MoveByOffset(0, 40).Click().SendKeys("SUMMARY TEST").Build().Perform();

                //save and quit
                Print(method, "save and quit");
                m.Click(By.Id("btnSave"));
                m.Click(By.Id("btnClose"));
                Print(method, "Finished");
            }
        }
        /**Used to add annotations on frmDocument window
         */
        private void AddAnnotations() {
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            m.Click(By.Id("lblType"));
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            action.MoveByOffset(424, -35).Build().Perform();
            Thread.Sleep(1000);
            action.Click().MoveByOffset(0, 100).Build().Perform();
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            action.ClickAndHold().MoveByOffset(0, 50).Build().Perform();
        }
        /** Searches intact for a keyword, if not found fails the test
         *  searchInput: string put in searchBar
         */
        #endregion

        public void Search(string searchInput) {
            method = MethodBase.GetCurrentMethod().Name;
            window = m.Locate(By.Id("frmIntactMain"));
            window = m.Locate(By.Name("radMenu1"), window);
            m.SendKeys(By.ClassName("WindowsForms10.EDIT.app.0.5c39d4"), searchInput, window);
            m.Click(By.Name("Search"), window);

            Thread.Sleep(1000);
            if(m.IsElementPresent(By.Name("Quick Search"))) {
                //m.Click(By.Name("OK"));
                Print(method, "Result not found");
                throw new AssertFailedException(method + ": Result Not Found");
            } else {
                Print(method, "Result Found");
            }
        }
        public void OpenOrganizer() {
            method = MethodBase.GetCurrentMethod().Name;
            Print(method, " Started");
            m.Click(By.Name("Organizer"));
            Print(method, " Finished");
        }
        /**
         * This method is going to add documents to batch review and then run through and both add a document to an existing document and attribute a new one.
         * Does usually run slow
         */
        public void BatchReview() {
            method = MethodBase.GetCurrentMethod().Name;
            AddDocsToCollector();

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
            AddDocBatchReview();

            m.Click(By.Id("btnClose"));
        }
        private void AddDocBatchReview() {
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
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            m.Click(By.Id("lblType"));
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            action.MoveByOffset(395, -37).Click().Click().Click().Click().Click().Build().Perform();
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);

            //custom fields
            Print(method, "custom fields");
            m.Click(By.Id("lblType"));
            action.MoveByOffset(150, 240).Click().SendKeys("1/1/2000").
                MoveByOffset(0, 20).Click().SendKeys("10").MoveByOffset(0, 20).Click().SendKeys("BATCH REVIEW ADD").Build().Perform();

            //edit fields
            Print(method, "edit fields ");
            m.Click(By.Id("lblType"));
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
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            m.Click(By.Id("lblType"));
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            action.MoveByOffset(375, -37).Click().Click().Build().Perform();
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);

            //custom fields
            Print(method, "custom fields");
            m.Click(By.Id("lblType"));
            action.MoveByOffset(150, 240).Click().SendKeys("1/1/2000").
                MoveByOffset(0, 20).Click().SendKeys("10").MoveByOffset(0, 20).Click().SendKeys("BATCH REVIEW ATTRIBUTION").Build().Perform();

            //edit fields
            Print(method, "edit fields ");
            m.Click(By.Id("lblType"));
            action.MoveByOffset(170, 80).Click().SendKeys("1/1/2000").MoveByOffset(0, 20).Click().SendKeys("BATCH AUTHOR TEST").
                MoveByOffset(0, 40).Click().SendKeys("BATCH ATTRIBUTING A DOCUMENT TEST").Build().Perform();

            m.Click(By.Id("btnSave"), window);
            m.Click(By.Id("btnClose"), window);
        }
        /**
         * Collects documents by all definition from InZone. Returns true if the definition recognized is the same name as a file in directory
         */
        public bool GetDocumentsFromInZone() {
            method = MethodBase.GetCurrentMethod().Name;
            AddDocsToCollector();
            window = m.Locate(By.Id("frmIntactMain"));
            window = m.Locate(By.Id("radMenu1"), window);
            m.Click(By.Name("&Intact"), window);
            window = m.Locate(By.Name("&Intact"), window);
            m.Click(By.Name("InZone"), window);
            Thread.Sleep(2000);
            window = m.Locate(By.Id("frmInZoneMain"));
            m.Click(By.Id("btnCollectScan"), window);
            Thread.Sleep(10000);
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
        /**
         * Method copies over files from one directory to another: Each time before InZone collects this is going to put files in the collector folder
         * Verify that the startPath always has files in it and those files shouldn't be removed from this folder when collected.
         */
        private void AddDocsToCollector() {
            method = MethodBase.GetCurrentMethod().Name;
            string startPath = ConfigurationManager.AppSettings.Get("InZoneStartPath");
            string endPath = ConfigurationManager.AppSettings.Get("InZoneCollectorPath");
            if (Directory.Exists(startPath) & Directory.Exists(endPath)) {
                foreach (string s in Directory.GetFiles(startPath)) {
                    File.Copy(s, Path.Combine(endPath, Path.GetFileName(s)), true);
                }
            } else {
                Print(method, "Starting or Ending path doesn't exist");
            }
        }
        private void AddRecognition() {
            m.Click(By.Name("Recognize"));
            window = m.Locate(By.Id("frmMainInteractive"));
            Thread.Sleep(1500);
            window = m.Locate(By.Id("btnSelect"), window);
            m.Click(By.Name("Select All"), window);
            m.Click(By.Id("btnRecgonize"));
            m.Click(By.Id("btnClose"));
        }
        /**UNTESTED
         * definitionName: definition of what you want to recognize
         * documentName: document name
         * input: string to search to see if the page recognizes it
         */
        public void TestRecognition(string definitionName, string documentName, string input) { // test to make sure there are documents in recognize
            CreateDocument();
            AddRecognition();
            OpenOrganizer();
            m.Click(By.Name(definitionName));
            window = m.Locate(By.Name(documentName));
            action.DoubleClick(window).Build().Perform();
            Thread.Sleep(2000);

            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            m.Click(By.Id("lblType"));
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);
            action.MoveByOffset(470, -40).Click().Build().Perform();
            Print(method, "x: " + Cursor.Position.X + " y: " + Cursor.Position.Y);

            window = m.Locate(By.Name("Find"));
            m.SendKeys(By.Name(""), input, m.Locate(By.Id("txtFind"), window));
            m.Click(By.Id("btnFind"));

            if(m.IsElementPresent(By.Name("Search Text Not Found"), window)) {
                throw new AssertFailedException("Search Text not found for Recognize");
            }
            m.Click(By.Id("btnCancel"));
            m.Click(By.Id("btnClose"));
        }
        //NOT FINISHED
        public void OpenUtil() {
            //Document indexing
            window = m.Locate(By.Name("&Administration"));
            window = m.Locate(By.Name("Utilities"), window);
            m.Click(By.Name("Index Documents..."), window);
            window = m.Locate(By.Id("frmDocumentIndexing"));
            m.Click(By.Id("btnFull"), window);
            m.Click(By.Id("btnClose"), window);

            //TODO:Fix expired documents another table
            window = m.Locate(By.Name("&Administration"));
            window = m.Locate(By.Name("Utilities"), window);
            m.Click(By.Name("View Expired Documents..."), window);
            m.Click(By.Id("Close"));

            //View licenses
            //window = m.Locate(By.Name("&Administration"));
            //window = m.Locate(By.Name("Utilities"), window);
            //m.Click(By.Name("View Licenses..."), window);
            //m.Click(By.Name("Close"));

            //Change Background image
            window = m.Locate(By.Name("&Administration"));
            window = m.Locate(By.Name("Utilities"), window);
            m.Click(By.Name("Change Background Image..."), window);
            Thread.Sleep(1500);
            window = m.Locate(By.Name("Select Client Background Image"), m.Locate(By.Id("frmIntactMain")));
            m.Click(By.Id("btnOK"), window);

            //Batch Recognize skip for now because of recognize test...

            //Settings Console
            window = m.Locate(By.Name("&Administration"));
            window = m.Locate(By.Name("Utilities"), window);
            m.Click(By.Name("Settings Console..."), window);
            m.Click(By.Id("&OK"));
            m.Click(By.Id("Close"));

            //Diagnostics Console
            window = m.Locate(By.Name("&Administration"));
            window = m.Locate(By.Name("Utilities"), window);
            m.Click(By.Name("Diagnostics Utility..."), window);
            m.Click(By.Id("Close"));

            //Refile Documents
            window = m.Locate(By.Name("&Administration"));
            window = m.Locate(By.Name("Utilities"), window);
            m.Click(By.Name("Refile Documents..."), window);
            m.Click(By.Id("Close"));

            //View Recognize Errors 
            window = m.Locate(By.Name("&Administration"));
            window = m.Locate(By.Name("Utilities"), window);
            m.Click(By.Name("View Recognize Errors..."), window);
            m.Click(By.Id("Close"));
        }

        public void AddToIPack() {
            window = m.Locate(By.Name("&Intact"), m.Locate(By.Name("radMenu1"))); 
            m.Click(By.Name("iPack"), window);
            m.Click(By.Id("rbnBatch"));
            Thread.Sleep(1000);
            if(m.IsElementPresent(By.Name("No documents were selected"))) {
                m.Click(By.Name("OK"));
            }
            Thread.Sleep(1000);

            m.Click(By.Id("radButton1")); 
            Thread.Sleep(5000);
        }

        public void AddToPortfolio() {
            window = m.Locate(By.Name("&Administration"), m.Locate(By.Name("radMenu1")));
            m.Click(By.Name("Portfolios"), window);
            window = m.Locate(By.Id("frmPortfolioMain"));
            m.Click(By.Id("btnGenerate"), window);
            //.....

            m.Click(By.Id("btnClose"));
            m.Click(By.Id("btnClose"));
        }

        #region 
        /** 
         * Found in test cleanup
         * saves screenshot in the directory specified, closes top window, returns path for the fail file
         */
        public string OnFail(string testName, string folderPath = "") {
            if (folderPath.Length < 2) {
                folderPath = ConfigurationManager.AppSettings.Get("AutomationScreenshots");
            }
            string path = Path.Combine(folderPath, testName + "_" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "-" + DateTime.Now.Year.ToString() +
                "_" + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + "-" + DateTime.Now.Second.ToString());
            ((ITakesScreenshot)driver).GetScreenshot().SaveAsFile(path, ImageFormat.Png);
                CloseWindow();
            return path;
        }
        /** 
         * Writes tests passed and failed in a file that can be set in config, appends to file and give the path to the screenshot
         */
        public void WriteFailFile(List<string> testsFailedNames, List<string> testsPassedNames) {
            method = MethodBase.GetCurrentMethod().Name;
            Print(method, "Started");
            using (StreamWriter file =
            new StreamWriter(ConfigurationManager.AppSettings.Get("TestFailedFile"), true)) {
                file.WriteLine(DateTime.Now.ToString() + "| " + testsFailedNames.Count.ToString() + " Tests failed | " + testsPassedNames.Count.ToString() + " Tests passed |");
                foreach (string name in testsFailedNames) {
                    file.WriteLine(name + " failed");
                }
                foreach (string name in testsPassedNames) {
                    file.WriteLine(name + " passed");
                }
                file.WriteLine(" ");
                file.WriteLine(" ");
            }
        }
        /**Take images from folder and put them on a word doc
         * Put at the end of tests.
         */
        public void WriteFailToWord() { //have to figure out how to delete things afterwards
            //method = MethodBase.GetCurrentMethod().Name;
            //Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            //Document wordDoc = word.Documents.Add();
            //Range range = wordDoc.Range();

            //string folderPath = ConfigurationManager.AppSettings.Get("AutomationScreenshots");
            //foreach (string image in Directory.GetFiles(folderPath)) {

            //    string imagePath = Path.Combine(folderPath, Path.GetFileName(image));
            //    Print(method, imagePath);

            //    InlineShape autoScaledInlineShape = range.InlineShapes.AddPicture(imagePath);
            //    float scaledWidth = autoScaledInlineShape.Width;
            //    float scaledHeight = autoScaledInlineShape.Height;
            //    autoScaledInlineShape.Delete();

            //    // Create a new Shape and fill it with the picture
            //    Shape newShape = wordDoc.Shapes.AddShape(1, 0, 0, scaledWidth, scaledHeight);
            //    newShape.Fill.UserPicture(imagePath);

            //    // Convert the Shape to an InlineShape and optional disable Border
            //    InlineShape finalInlineShape = newShape.ConvertToInlineShape();
            //    finalInlineShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            //    // Cut the range of the InlineShape to clipboard
            //    finalInlineShape.Range.Cut();

            //    // And paste it to the target Range
            //    range.Paste();
            //}
            //wordDoc.SaveAs2(ConfigurationManager.AppSettings.Get("TestFailedDocs"));
            //word.Quit();
        }
        #endregion

        #region
        public void CloseDriver() {
            driver.Quit();
            Print(method, "DRIVER CLOSED");
        }
        public void CloseWindow() {
            if (m.IsElementPresent(By.Id("btnClose"))) {
                m.Click(By.Id("btnClose"));
            }
            else if (m.IsElementPresent(By.Id("Close"))) {
                m.Click(By.Id("Close"));
            } else {
                action.KeyDown(OpenQA.Selenium.Keys.Alt).SendKeys(OpenQA.Selenium.Keys.F4).KeyUp(OpenQA.Selenium.Keys.Alt).Build().Perform();
            }
            Print(method, " window closed");
        }
        #endregion

        private void Print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
    }
}