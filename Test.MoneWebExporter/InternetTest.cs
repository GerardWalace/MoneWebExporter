using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MoneWebExporter.Excel;
using System.IO;
using MoneWebExporter.Data;
using System.Collections.Generic;
using MoneWebExporter.Internet;

namespace Test.MoneWebExporter
{
    [TestClass]
    public class InternetTest
    {
        const string TEST_ADRESSE = "http://sodexo-steria.moneweb.fr/";
        const string TEST_LOGIN = "DECKERM";
        const string TEST_PASSWORD = "s0d3x023";

        readonly Parametres Parametres = new Parametres()
        {
            AdresseMoneWeb = TEST_ADRESSE,
            Login = TEST_LOGIN,
            MotDePasse = TEST_PASSWORD,
        };

        [TestMethod]
        public void TestConnexionSite()
        {
            InternetBrowser browser = new InternetBrowser(Parametres);

            bool connexion = browser.ConnexionSite();

            Assert.IsTrue(connexion);
        }

        [TestMethod]
        public void TestTelechargerTicketsPage()
        {
            InternetBrowser browser = new InternetBrowser(Parametres);

            bool connexion = browser.ConnexionSite();
            List<Ticket> tickets = browser.TelechargerTicketsPage();

            Assert.IsNotNull(tickets);
            Assert.IsTrue(tickets.Any());
        }

        [TestMethod]
        public void TestPageSuivante()
        {
            InternetBrowser browser = new InternetBrowser(Parametres);

            bool connexion = browser.ConnexionSite();
            bool pageSuivante = browser.PageSuivante();

            Assert.IsTrue(pageSuivante);
        }

        [TestMethod]
        public void TestTelechargerAllTickets()
        {
            InternetBrowser browser = new InternetBrowser(Parametres);

            List<Ticket> tickets = browser.TelechargerAllTickets();

            Assert.IsNotNull(tickets);
            Assert.IsTrue(tickets.Any());
        }
    }
}
