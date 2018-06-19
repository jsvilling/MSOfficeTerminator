using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeKiller.Killers.OfficeApplicationKiller;

namespace OfficeKillerTest.IT
{
    [TestClass]
    public abstract class AppKillerTest<A> 
    {
        OfficeApplicationKiller<A> appKiller;

        [TestMethod]
        public void TestAppKiller_AppLaunched()
        {
            // Given
            appKiller = initAppKiller();
            LaunchApp();

            // When
            appKiller.Kill();

            // Then
            Process[] remainingAppProcesses = Process.GetProcessesByName(appKiller.ProcessName);
            Assert.AreEqual(0, remainingAppProcesses.Length);
        }

        [TestMethod]
        public void TestAppKiller_NoAppLaunched()
        {
            // Given
            appKiller = initAppKiller();

            // When
            appKiller.Kill();

            // Then
            Process[] remainingWordProcesses = Process.GetProcessesByName(appKiller.ProcessName);
            Assert.AreEqual(0, remainingWordProcesses.Length);
        }

        protected abstract OfficeApplicationKiller<A> initAppKiller();

        protected abstract void LaunchApp();
    }
}
