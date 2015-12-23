using HtmlAgilityPack;
using MoneWebExporter.Data;
using MoneWebExporter.Excel;
using MoneWebExporter.Internet;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using (ExcelBrowser xlBrowser = new ExcelBrowser(ExcelBrowser.FILE_NAME))
                {
                    //Si le fichier n'existe pas, le créer, afficher un message à l'utilisateur, quitter l'application
                    if (!xlBrowser.ExistenceFichier())
                    {
                        xlBrowser.CreerFichier();
                        xlBrowser.Sauvegarder();
                        Console.WriteLine(String.Format("Le fichier Excel {0} vient d'être créé.", ExcelBrowser.FILE_NAME));
                        Console.WriteLine(String.Format("Merci de remplir l'onglet {0} avant de relancer l'application.", ExcelBrowser.SHEETNAME_PARAMETRES));
                    }
                    else
                    {
                        //Si le fichier existe on l'ouvre
                        xlBrowser.OuvrirFichier();

                        //-	Chargement onglet paramètre
                        Parametres parametres = xlBrowser.ChargerParametres();
                        //-	Chargement onglet historique
                        List<LigneLog> historiques = xlBrowser.ChargerHistorique();
                        //-	Chargement onglet Repas
                        List<Ticket> tickets = xlBrowser.ChargerTickets();
                        int nbTicketAvant = tickets.Count;

                        try
                        {
                            //Si les paramètres ne sont pas valides, afficher un message à l'utilisateur puis quitter l'application
                            if (String.IsNullOrWhiteSpace(parametres.AdresseMoneWeb))
                            {
                                Console.WriteLine("La saisie d'une adresse de site *.moneweb.fr dans le fichier Excel est obligatoire.");
                            }
                            else if (String.IsNullOrWhiteSpace(parametres.Login))
                            {
                                Console.WriteLine("La saisie d'un login dans le fichier Excel est obligatoire.");
                            }
                            else if (String.IsNullOrWhiteSpace(parametres.MotDePasse))
                            {
                                Console.WriteLine("La saisie d'un mot de passe dans le fichier Excel est obligatoire.");
                            }
                            else
                            {
                                //Connexion au site web, téléchargement des repas
                                InternetBrowser intBrowser = new InternetBrowser(parametres);
                                tickets.AddRange(intBrowser.TelechargerAllTickets());
                                tickets = tickets.Distinct().ToList();
                                int nbTicketApres = tickets.Count;

                                //On effectue la mesure du nombre de repas, on log un lancement reussi dans le fichier excel, on enregistre les repas, on quitte en affichant le nombre de repas.
                                int nbRepas = CaculerNombreRepas(parametres.DerniereReleve, tickets);
                                Console.WriteLine(String.Format("Vous avez pris {0} Repas depuis le {1}", nbRepas, parametres.DerniereReleve));
                                Console.WriteLine("Ce nombre est à déduire du nombre de repas auquels vous aviez droit.");

                                DateTime? dateMax = tickets.Max(t => DateHelper.FromString(t.Date));
                                historiques.Add(CreerLigneLog(dateMax, parametres.DerniereReleve, nbTicketApres - nbTicketApres, nbRepas));
                                parametres.DerniereReleve = dateMax.ToString();

                                // On enregistre tout ça
                                xlBrowser.EnregistrerParametres(parametres);
                                xlBrowser.EnregistrerHistorique(historiques);
                                xlBrowser.EnregistrerTickets(tickets);
                                xlBrowser.Sauvegarder();
                            }
                        }
                        catch(Exception e)
                        {
                            historiques.Add(CreerLigneErreur(e));
                            xlBrowser.EnregistrerHistorique(historiques);
                            xlBrowser.Sauvegarder();
                        }
                    }
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("Erreur !");
                Console.WriteLine("L'erreur suivante s'est produite et n'a pas pu être loggée :");
                Console.WriteLine(e.Message);
            }
        }

        private static LigneLog CreerLigneErreur(Exception e)
        {
            //TODO
            return new LigneLog();
        }

        private static LigneLog CreerLigneLog(DateTime? dateMax, string p1, int p2, int nbRepas)
        {
            //TODO
            return new LigneLog();
        }

        private static int CaculerNombreRepas(string p, List<Ticket> tickets)
        {
            //TODO
            return tickets.Count;
        }

        
    }
}
