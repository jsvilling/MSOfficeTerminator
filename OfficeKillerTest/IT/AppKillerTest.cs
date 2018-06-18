using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeKiller.Killers.OfficeApplicationKiller;

namespace OfficeKillerTest.IT
{
    [TestClass]
    public abstract class AppKillerTest
    {
        [TestInitialize]
        public void TestInitialize()
        {
            if (IsInstanceRunning())
            {
                Console.Write("Please quit all Office Applications before running any tests");
                throw new SystemException("unexpected office app running");
            }
        }

        protected abstract bool IsInstanceRunning();
    }
}
