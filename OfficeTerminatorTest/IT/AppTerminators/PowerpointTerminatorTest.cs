using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTerminator.Terminators.OfficeApplicationTerminator;

namespace OfficeTerminatorTest.IT
{
    [TestClass]
    public class PowerpointTerminatorTest : AppTerminatorTest<Application>
    {
        protected override OfficeApplicationTerminator<Application> initAppTerminator()
        {
            return new PowerpointTerminator();
        }

        protected override void LaunchApp()
        {
            Application ppt = new Application();
            ppt.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            ppt.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
        }
    }
}
