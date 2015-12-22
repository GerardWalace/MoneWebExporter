using MoneWebExporter.Data;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter.Excel
{
    public class ExcelBrowser : IDisposable
    {
        public const string FILE_NAME = "MoneWebExporter.xlsx";
        private const string SHEETNAME_PARAMETRES = "PARAMETRES";
        private const string SHEETNAME_HISTORIQUES = "HISTORIQUES";
        private const string SHEETNAME_TICKETS = "TICKETS";

        private FileInfo xlFichier;
        private ExcelPackage xlPackage;


        public ExcelBrowser(string pathFichierExcel)
        {
            if (!String.IsNullOrWhiteSpace(pathFichierExcel))
                xlFichier = new FileInfo(pathFichierExcel);            
        }

        public bool ExistenceFichier()
        {
            return xlFichier.Exists;
        }

        public void CreerFichier()
        {
            if (xlFichier != null)
                xlPackage = new ExcelPackage(xlFichier);
            else
                xlPackage = new ExcelPackage();


            EnregistrerParametres(new Parametres());
            EnregistrerHistorique(new List<LigneLog>());
            EnregistrerTickets(new List<Ticket>());
        }

        public void OuvrirFichier()
        {
            if (!ExistenceFichier())
            {
                throw new FireException("Le fichier xlsx ne peut pas être ouvert car il n'existe pas.");
            }
            else
            {
                xlPackage = new ExcelPackage(xlFichier);
            }
        }

        private ExcelWorksheet OpenClearSheet(string name)
        {
            ExcelWorksheet worksheet;
            worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault(w => w.Name == SHEETNAME_PARAMETRES);
            if (worksheet != null)
                xlPackage.Workbook.Worksheets.Delete(worksheet);

            worksheet = xlPackage.Workbook.Worksheets.Add(SHEETNAME_PARAMETRES);

            return worksheet;
        }

        private ExcelWorksheet OpenSheet(string name)
        {
            ExcelWorksheet worksheet;
            worksheet = xlPackage.Workbook.Worksheets.FirstOrDefault(w => w.Name == SHEETNAME_PARAMETRES);
            if (worksheet == null)
                throw new FireException("Le fichier xlsx n'est pas valide. Il ne contient pas l'onglet des parametres.");

            return worksheet;
        }

        public void EnregistrerParametres(Parametres parametres)
        {
            ExcelWorksheet worksheet = OpenClearSheet(SHEETNAME_PARAMETRES);

            worksheet.Cells[1, 1].Value = "Adresse Internet MoneWeb";
            worksheet.Cells[1, 2].Value = parametres.AdresseMoneWeb;

            worksheet.Cells[2, 1].Value = "Login MoneWeb";
            worksheet.Cells[2, 2].Value = parametres.Login;

            worksheet.Cells[3, 1].Value = "Password MoneWeb";
            worksheet.Cells[3, 2].Value = parametres.MotDePasse;

            worksheet.Cells[4, 1].Value = "Date dernière relève";
            worksheet.Cells[4, 2].Value = parametres.DerniereReleve;
        }

        public Parametres ChargerParametres()
        {
            ExcelWorksheet worksheet = OpenSheet(SHEETNAME_PARAMETRES);
            Parametres parametres = new Parametres();

            parametres.AdresseMoneWeb = worksheet.Cells[1, 2].Value.ToString();
            parametres.Login = worksheet.Cells[2, 2].Value.ToString();
            parametres.MotDePasse = worksheet.Cells[3, 2].Value.ToString();
            parametres.DerniereReleve = worksheet.Cells[4, 2].Value.ToString();

            return parametres;
        }

        public void EnregistrerHistorique(List<LigneLog> historique)
        {
            ExcelWorksheet worksheet = OpenClearSheet(SHEETNAME_HISTORIQUES);

            for (int i = 0; i < historique.Count;i++)
            {
                int j = 1;
                worksheet.Cells[2 + i, j++].Value = historique[i].DateLancement;
                worksheet.Cells[2 + i, j++].Value = historique[i].NombreNouvellesLignes;
                worksheet.Cells[2 + i, j++].Value = historique[i].DateDerniereReleveUtilisee;
                worksheet.Cells[2 + i, j++].Value = historique[i].NombreSubventions;
                worksheet.Cells[2 + i, j++].Value = historique[i].MessageErreur;
                worksheet.Cells[2 + i, j++].Value = historique[i].PileAppelErreur;
            }
        }

        public List<LigneLog> ChargerHistorique()
        {
            ExcelWorksheet worksheet = OpenSheet(SHEETNAME_HISTORIQUES);
            List<LigneLog> historiques = new List<LigneLog>();

            int i = 0;
            while (worksheet.Cells[2 + i, 1].Value != null)
            {
                LigneLog historique = new LigneLog();
                int j = 1;

                historique.DateLancement = worksheet.Cells[2 + i, j++].Value.ToString();
                historique.NombreNouvellesLignes = int.Parse(worksheet.Cells[2 + i, j++].Value.ToString());
                historique.DateDerniereReleveUtilisee = worksheet.Cells[2 + i, j++].Value.ToString();
                historique.NombreSubventions = int.Parse(worksheet.Cells[2 + i, j++].Value.ToString());
                historique.MessageErreur = worksheet.Cells[2 + i, j++].Value.ToString();
                historique.PileAppelErreur = worksheet.Cells[2 + i, j++].Value.ToString();

                historiques.Add(historique);
                i++;
            }

            return historiques;
        }

        public void EnregistrerTickets(List<Ticket> tickets)
        {
            ExcelWorksheet worksheet = OpenClearSheet(SHEETNAME_TICKETS);

            for (int i = 0; i < tickets.Count; i++)
            {
                int j = 1;
                worksheet.Cells[2 + i, j++].Value = tickets[i].Date;
                worksheet.Cells[2 + i, j++].Value = tickets[i].Activite;
                worksheet.Cells[2 + i, j++].Value = tickets[i].AncienSolde;
                worksheet.Cells[2 + i, j++].Value = tickets[i].Plateau;
                worksheet.Cells[2 + i, j++].Value = tickets[i].Financier;
                worksheet.Cells[2 + i, j++].Value = tickets[i].NouveauSolde;
            }
        }

        public List<Ticket> ChargerTickets()
        {
            ExcelWorksheet worksheet = OpenSheet(SHEETNAME_TICKETS);
            List<Ticket> tickets = new List<Ticket>();

            int i = 0;
            while (worksheet.Cells[2 + i, 1].Value != null)
            {
                Ticket ticket = new Ticket();
                int j = 1;

                ticket.Date = worksheet.Cells[2 + i, j++].Value.ToString();
                ticket.Activite = worksheet.Cells[2 + i, j++].Value.ToString();
                ticket.AncienSolde = worksheet.Cells[2 + i, j++].Value.ToString();
                ticket.Plateau = worksheet.Cells[2 + i, j++].Value.ToString();
                ticket.Financier = worksheet.Cells[2 + i, j++].Value.ToString();
                ticket.NouveauSolde = worksheet.Cells[2 + i, j++].Value.ToString();

                tickets.Add(ticket);
                i++;
            }

            return tickets;
        }

        public void Sauvegarder()
        {
            if (xlPackage != null)
                xlPackage.Save();
        }

        public void Dispose()
        {
            if (xlPackage != null)
                xlPackage.Dispose();
        }
    }
}
