using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Shared.Tests
{
    [TestClass]
    public class SharedTests
    {
        [TestMethod]
        public void GetVersion_WhenRunningProcessIsNotInOfficeVersionDictionary_ReturnsExpectedResult()
        {
            var result = Shared.GetVersion();
        }
    }
}
