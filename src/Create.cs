using log4net;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace WiniumTests.src {
    class Create {
        IWebElement window;
        readonly WiniumMethods m;
        Actions action;
        string method = "";
        ILog debugLog;

        public Create(WiniumMethods m, Actions action, ILog debugLog) {
            this.m = m;
            this.action = action;
            this.debugLog = debugLog; 
        }
        /**This is going to a specified amount of definitions with random name for each blank.
       */
        public void CreateNewDefinition(int? numberOfDefinitions = 1, string definitionName = "") {
            method = MethodBase.GetCurrentMethod().Name;
            Print(method, "Started");
            //check if maximized
            window = m.Locate(By.Id("frmIntactMain"));
            if (m.IsElementPresent(By.Name("Maximize"), window)) {
                m.Click(By.Name("Maximize"), window);
            }
            window = m.Locate(By.Name("radMenu1"), window);
            m.Click(By.Name("&Administration"), window);
            window = m.Locate(By.Name("&Administration"), window);
            m.Click(By.Name("Definitions"), window);

            if (definitionName.Length < 2) {
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
                        try { element.SendKeys(definitionName + " " + num); } catch (Exception) { }
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
            //check if maximized
            window = m.Locate(By.Id("frmIntactMain"));
            if (m.IsElementPresent(By.Name("Maximize"), window)) {
                m.Click(By.Name("Maximize"), window);
            }
            window = m.Locate(By.Name("radMenu1"), window);
            m.Click(By.Name("&Administration"), window);

            m.Click(By.Name("Types"), m.Locate(By.Name("&Administration")));

            if (typeName.Length < 2) {
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
            //check if maximized
            window = m.Locate(By.Id("frmIntactMain"));
            if (m.IsElementPresent(By.Name("Maximize"), window)) {
                m.Click(By.Name("Maximize"), window);
            }

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
        public void Print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
    }
}
