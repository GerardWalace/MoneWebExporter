using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MoneWebExporter.Excel;
using System.IO;
using MoneWebExporter.Data;
using System.Collections.Generic;

namespace Test.MoneWebExporter
{
    [TestClass]
    public class DataTest
    {
        [TestMethod]
        public void TestTicketEquals()
        {
            Ticket a = null;
            Ticket b = null;

            Assert.IsTrue(a == b);

            b = new Ticket() { Date = DateTime.Now.ToString() };
            Assert.IsFalse(a == b);

            a = new Ticket() { Date = DateTime.Now.AddDays(1).ToString() };
            Assert.IsFalse(a == b);

            a.Date = b.Date;
            Assert.IsTrue(a == b);
        }

        [TestMethod]
        public void TestDistinctTickets()
        {
            DateTime now = DateTime.Now;
            List<Ticket> tickets = new List<Ticket>()
            {
                new Ticket(){ Date = now.ToString() },
                new Ticket(){ Date = now.ToString() },
                new Ticket(){ Date = now.AddDays(1).ToString() },
                new Ticket(){ Date = now.ToString() },
                new Ticket(){ Date = now.AddDays(2).ToString() },
                new Ticket(){ Date = now.ToString() },
            };

            Assert.IsTrue(tickets.Count == 6);

            tickets = tickets.Distinct().ToList();
            Assert.IsTrue(tickets.Count == 3);
        }
    }
}
