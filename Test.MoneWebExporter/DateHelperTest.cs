using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MoneWebExporter;

namespace Test.MoneWebExporter
{
    [TestClass]
    public class DateHelperTest
    {
        [TestMethod]
        public void TestLoop()
        {
            DateTime? date = DateTime.Now;
            string str = DateHelper.ToString(date);
            DateTime? date2 = DateHelper.FromString(str);
                
            Assert.IsTrue(date2.HasValue);
            Assert.AreEqual(date.ToString(), date2.ToString());
            Assert.AreEqual(date.Value.Second, date2.Value.Second);
        }
    }
}
