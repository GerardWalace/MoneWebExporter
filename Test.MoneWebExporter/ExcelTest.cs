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
    public class ExcelTest
    {
        [TestMethod]
        public void TestCreerFichier()
        {
            using (ExcelBrowser browser = new ExcelBrowser(ExcelBrowser.FILE_NAME))
            {
                browser.CreerFichier();
                browser.Sauvegarder();
            }

            Assert.IsTrue(File.Exists(ExcelBrowser.FILE_NAME));
        }

        [TestMethod]
        public void TestLoopParametres()
        {
            Parametres parametres1 = new Parametres()
                {
                    AdresseMoneWeb = "http://www.test.fr/",
                    Login = "MonLogin",
                    MotDePasse = "MonMotDePasse",
                    DerniereReleve = "12/12/2012 12:12:12"
                };
            Parametres parametres2 = null;

            using (ExcelBrowser browser = new ExcelBrowser(""))
            {
                browser.CreerFichier();
                browser.EnregistrerParametres(parametres1);
                parametres2 = browser.ChargerParametres();
            }

            Assert.IsNotNull(parametres2);
            Assert.AreEqual(parametres1.AdresseMoneWeb, parametres2.AdresseMoneWeb);
            Assert.AreEqual(parametres1.Login, parametres2.Login);
            Assert.AreEqual(parametres1.MotDePasse, parametres2.MotDePasse);
            Assert.AreEqual(parametres1.DerniereReleve, parametres2.DerniereReleve);
        }

        [TestMethod]
        public void TestLoopHistoriques()
        {
            LigneLog historique11 = new LigneLog()
            {
                DateLancement = "12/12/2012",
                NombreNouvellesLignes = 4,
                DateDerniereReleveUtilisee = "11/12/2012",
                NombreSubventions = 5,
                MessageErreur = "Test",
                PileAppelErreur = "Truc",
            };
            List<LigneLog> historiques1 = new List<LigneLog>()
            {
                historique11,
                historique11,
            };
            LigneLog historique21 = null;
            List<LigneLog> historiques2 = null;

            using (ExcelBrowser browser = new ExcelBrowser(""))
            {
                browser.CreerFichier();
                browser.EnregistrerHistorique(historiques1);
                historiques2 = browser.ChargerHistorique();
            }

            Assert.IsNotNull(historiques2);
            Assert.AreEqual(historiques1.Count, historiques2.Count);
            historique21 = historiques2.First();
            Assert.AreEqual(historique11.DateLancement, historique21.DateLancement);
            Assert.AreEqual(historique11.NombreNouvellesLignes, historique21.NombreNouvellesLignes);
            Assert.AreEqual(historique11.DateDerniereReleveUtilisee, historique21.DateDerniereReleveUtilisee);
            Assert.AreEqual(historique11.NombreSubventions, historique21.NombreSubventions);
            Assert.AreEqual(historique11.MessageErreur, historique21.MessageErreur);
            Assert.AreEqual(historique11.PileAppelErreur, historique21.PileAppelErreur);
        }

        [TestMethod]
        public void TestLoopTickets()
        {
            Ticket ticket11 = new Ticket()
            {
                Date = "12/12/2012",
                Activite = "SODEXO",
                AncienSolde = "12 €",
                Plateau = "13 €",
                Financier = "14 €",
                NouveauSolde = "15 €",
            };
            List<Ticket> tickets1 = new List<Ticket>()
            {
                ticket11,
                ticket11,
            };
            Ticket ticket21 = null;
            List<Ticket> tickets2 = null;

            using (ExcelBrowser browser = new ExcelBrowser(""))
            {
                browser.CreerFichier();
                browser.EnregistrerTickets(tickets1);
                tickets2 = browser.ChargerTickets();
            }

            Assert.IsNotNull(tickets2);
            Assert.AreEqual(tickets1.Count, tickets2.Count);
            ticket21 = tickets2.First();
            Assert.AreEqual(ticket11.Date, ticket21.Date);
            Assert.AreEqual(ticket11.Activite, ticket21.Activite);
            Assert.AreEqual(ticket11.AncienSolde, ticket21.AncienSolde);
            Assert.AreEqual(ticket11.Plateau, ticket21.Plateau);
            Assert.AreEqual(ticket11.Financier, ticket21.Financier);
            Assert.AreEqual(ticket11.NouveauSolde, ticket21.NouveauSolde);
        }
    }
}
