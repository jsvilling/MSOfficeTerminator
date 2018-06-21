using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTerminator.Terminators.OfficeApplicationTerminator;
using Microsoft.Office.Interop.Word;
using OfficeTerminatorTest.IT;
using System.Diagnostics;

namespace OfficeTerminatorTest 
{
    [TestClass]
    public class WordTerminatorTest : AppTerminatorTest<Application>
    {
        protected override OfficeApplicationTerminator<Application> initAppTerminator()
        {
            return new WordTerminator();
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