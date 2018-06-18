using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeKiller.Killers.OfficeApplicationKiller;
using Microsoft.Office.Interop.Word;

namespace OfficeKillerTest
{
    [TestClass]
    public class WordKillerTest
    {

        [TestInitialize]
        public void TestInitialize()
        {
            // assert that no word processes are running
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            // clean up any remaining word processes
        }

        [TestMethod]
        public void TestWordKiller_HappyPath()
        {
            // Given
            // spawn word process
            OfficeApplicationKiller appKiller = new WordKiller();

            // When
            appKiller.Kill();

            // Then
            Assert.IsFalse(IsWordInstanceRunning());
        }

        private void LaunchWord()
        {
            Application wordApp = new Application();
            wordApp.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApp.Application.ScreenUpdating = false;
            wordApp.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            wordApp.Application.Visible = false;
        }

        private bool IsWordInstanceRunning()
        {
            return (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application") != null;
        }
    }
}
