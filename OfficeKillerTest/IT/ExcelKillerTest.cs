using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeKiller.Killers.OfficeApplicationKiller;
using OfficeKillerTest.IT;

namespace OfficeKillerTest
{
    [TestClass]
    public class ExcelKillerTest : AppKillerTest<Application>
    {
        protected override OfficeApplicationKiller<Application> initAppKiller()
        {
            return new ExcelKiller();
        }

        protected override void LaunchApp()
        {
            Application excelApp = new Application();
            excelApp.Application.DisplayAlerts = false;
            excelApp.Application.ScreenUpdating = false;
            excelApp.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            excelApp.Application.Visible = false;
        }
    }
}