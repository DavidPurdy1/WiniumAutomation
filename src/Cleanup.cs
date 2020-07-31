using log4net;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;

namespace WiniumTests.src {
    /// <summary>
    /// Class containing methods used at the end of a test case or test run.
    /// </summary>
    public class Cleanup {
        readonly WiniumMethods m;
        string method = "";
        readonly ILog debugLog;

        public Cleanup(WiniumMethods m, ILog debugLog) {
            this.m = m;
            this.debugLog = debugLog;
        }
        /** 
         * Found in test cleanup
         * saves screenshot in the directory specified, closes top window, returns path for the fail file
         */
        public string OnFail(string testName, string folderPath = "") {
            if (folderPath.Length < 2) {
                folderPath = ConfigurationManager.AppSettings.Get("AutomationScreenshots");
            }

            //YYYY-MM-DD__HH-MM-SS
            string dateAndTime = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "__"
                + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + "-" + DateTime.Now.Second.ToString();

            //creates file, stores screenshot in path
            string path = Path.Combine(folderPath, testName + "_" + dateAndTime);
            m.GetScreenshot().SaveAsFile(path + ".PNG", ImageFormat.Png);

            //CloseWindow();

            return path;
        }
        /** 
         * Writes tests passed and failed in a file that can be set in config, appends to file and give the path to the screenshot
         */
        public void WriteFailFile(List<string> testsFailedNames, List<string> testsPassedNames, List<string> imagePaths) {
            method = MethodBase.GetCurrentMethod().Name;
            Print(method, "Started");

            //YYYY-MM-DD__HH-MM-SS
            string dateAndTime = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString() + "__"
                + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + "-" + DateTime.Now.Second.ToString();

            //writes the result file information
            using (StreamWriter file =
            new StreamWriter(ConfigurationManager.AppSettings.Get("TestFailedFile") + dateAndTime + ".txt", false)) {
                var versionInfo = FileVersionInfo.GetVersionInfo(ConfigurationManager.AppSettings.Get("IntactPath"));
                string version = versionInfo.FileVersion;
                file.WriteLine("");
                file.WriteLine(DateTime.Now.ToString());
                file.WriteLine("Tests failed| " + testsFailedNames.Count.ToString());
                file.WriteLine("Tests passed| " + testsPassedNames.Count.ToString());
                file.WriteLine("Tester: " + System.Security.Principal.WindowsIdentity.GetCurrent().Name);
                file.WriteLine("App Version: " + version);
                file.WriteLine("App Name: " + ConfigurationManager.AppSettings.Get("IntactPath"));

                //tests failed and passed on New Line
                int i = 0;
                foreach (string name in testsFailedNames) {
                    file.WriteLine("failed| " + name);
                    file.WriteLine(imagePaths[i]);
                    i++;
                }
                foreach (string name in testsPassedNames) {
                    file.WriteLine("passed| " + name);
                }
            }
        }
        /**Take images from folder and put them on a word doc
         * Put at the end of tests.
         */
        public void WriteFailToWord() {
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
        public void SendToDB() {
            string connectionString = ConfigurationManager.AppSettings.Get("DBConnection");
            DataExporter exporter = new DataExporter(connectionString);
            exporter.ParseFile(new TestData());
        }

        #region Close
        public void CloseDriver() {
            m.CloseDriver();
        }
        #endregion

        public void Print(string method, string toPrint) {
            debugLog.Info(method + " " + toPrint);
        }
    }
}
