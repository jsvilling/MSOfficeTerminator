using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeKiller.Killers.OfficeApplicationKiller;
using Microsoft.Office.Interop.Word;
using OfficeKillerTest.IT;
using System.Diagnostics;

namespace OfficeKillerTest 
{
    [TestClass]
    public class WordKillerTest : AppKillerTest<Application>
    {
        protected override OfficeApplicationKiller<Application> initAppKiller()
        {
            return new WordKiller();
        }

        protected override void LaunchApp()
        {
            Application wordApp = new Application();
            wordApp.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            wordApp.Application.ScreenUpdating = false;
            wordApp.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            wordApp.Application.Visible = false;
        }
    }
}