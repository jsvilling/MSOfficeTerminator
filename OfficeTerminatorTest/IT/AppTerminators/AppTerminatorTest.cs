using System;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTerminator.Terminators.OfficeApplicationTerminator;

namespace OfficeTerminatorTest.IT
{
    [TestClass]
    public abstract class AppTerminatorTest<A> 
    {
        OfficeApplicationTerminator<A> appTerminator;

        [TestMethod]
        public void TestAppTerminator_AppLaunched()
        {
            // Given
            appTerminator = initAppTerminator();
            LaunchApp();

            // When
            appTerminator.Terminate();

            // Then
            Process[] remainingAppProcesses = Process.GetProcessesByName(appTerminator.ProcessName);
            Assert.AreEqual(0, remainingAppProcesses.Length);
        }

        [TestMethod]
        public void TestAppTerminator_NoAppLaunched()
        {
            // Given
            appTerminator = initAppTerminator();

            // When
            appTerminator.Terminate();

            // Then
            Process[] remainingWordProcesses = Process.GetProcessesByName(appTerminator.ProcessName);
            Assert.AreEqual(0, remainingWordProcesses.Length);
        }

        protected abstract OfficeApplicationTerminator<A> initAppTerminator();

        protected abstract void LaunchApp();
    }
}
