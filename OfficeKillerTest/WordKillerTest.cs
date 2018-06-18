using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeKiller.Killers.OfficeApplicationKiller;

namespace OfficeKillerTest
{
    [TestClass]
    public class WordKillerTest
    {

        [TestInitialize]
        public void TestInitialize()
        {
            // assert that no word processes are running
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            // clean up any remaining word processes
        }

        [TestMethod]
        public void TestWordKiller_HappyPath()
        {
            // Given
            // spawn word process
            // init wordkiller

            // When
            // invoke wordkiller.Kill()

            // Then
            // assert that no more word processes are running
        }
    }
}
