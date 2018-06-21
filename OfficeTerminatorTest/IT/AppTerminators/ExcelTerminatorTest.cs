using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTerminator.Terminators.OfficeApplicationTerminator;
using OfficeTerminatorTest.IT;

namespace OfficeTerminatorTest
{
    [TestClass]
    public class ExcelTerminatorTest : AppTerminatorTest<Application>
    {
        protected override OfficeApplicationTerminator<Application> initAppTerminator()
        {
            return new ExcelTerminator();
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