using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeKiller.Killers.OfficeApplicationKiller;
using Microsoft.Office.Interop.Word;
using OfficeKillerTest.IT;
using System.Diagnostics;

namespace OfficeKillerTest
{
    [TestClass]
    public class WordKillerTest
    {

        OfficeApplicationKiller<Application> appKiller = new WordKiller();

        [TestMethod]
        public void TestWordKiller_HappyPath()
        {
            // Given
            LaunchWord();

            // When
            appKiller.Kill();

            // Then
            Process[] remainingWordProcesses = Process.GetProcessesByName(appKiller.ProcessName);
            Assert.AreEqual(0, remainingWordProcesses.Length);
        }

        private void LaunchWord()
        {
            Application wordApp = new Application();
            wordApp.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApp.Application.ScreenUpdating = false;
            wordApp.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            wordApp.Application.Visible = false;
        }
    }
}
