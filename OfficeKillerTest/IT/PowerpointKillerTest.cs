using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeKiller.Killers.OfficeApplicationKiller;

namespace OfficeKillerTest.IT
{
    [TestClass]
    public class PowerpointKillerTest : AppKillerTest<Application>
    {
        protected override OfficeApplicationKiller<Application> initAppKiller()
        {
            return new PowerpointKiller();
        }

        protected override void LaunchApp()
        {
            Application ppt = new Application();
            ppt.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            ppt.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
        }
    }
}
