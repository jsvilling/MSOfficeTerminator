//using System;
//using Microsoft.Office.Interop.Excel;
//using Microsoft.VisualStudio.TestTools.UnitTesting;
//using OfficeKiller.Killers.OfficeApplicationKiller;
//using OfficeKillerTest.IT;

//namespace OfficeKillerTest
//{
//    [TestClass]
//    public class ExcelKillerTest : AppKillerTest
//    {

//        [TestMethod]
//        public void TestExcelKiller_HappyPath()
//        {
//            // Given
//            LaunchExcel();
//            OfficeApplicationKiller<Application> appKiller = new ExcelKiller();

//            // When
//            appKiller.Kill();

//            // Then
//            Assert.IsFalse(IsInstanceRunning());
//        }

//        protected override bool IsInstanceRunning()
//        {
//            Application app = InstanceUtils.FindRunningInstance<Application>("Excel.Application");
//            return app != null;
//        }

//        protected override void ShutdownRemainingInstances()
//        {
//            InstanceUtils.FindRunningInstance<Application>("Excel.Application").Quit();
//        }

//        private void LaunchExcel()
//        {
//            Application excelApp = new Application();
//            excelApp.Application.DisplayAlerts = false;
//            excelApp.Application.ScreenUpdating = false;
//            excelApp.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
//            excelApp.Application.Visible = false;
//        }
//    }
//}
