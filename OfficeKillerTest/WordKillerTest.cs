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
            if (IsWordInstanceRunning())
            {
                Console.Write("Please quit all Office Applications before running any tests");
                throw new SystemException("unexpected office app running");
            }
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            if (IsWordInstanceRunning())
            {
                InstanceUtils.FindRunningInstance<Application>("Word.Application").Quit();
            }
        }

        [TestMethod]
        public void TestWordKiller_HappyPath()
        {
            // Given
            LaunchWord();
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
            Application app = InstanceUtils.FindRunningInstance<Application>("Word.Application");
            return  app != null;
        }
    }
}
