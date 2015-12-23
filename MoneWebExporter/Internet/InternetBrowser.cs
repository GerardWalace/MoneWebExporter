using HtmlAgilityPack;
using MoneWebExporter.Data;
using MoneWebExporter.Internet.HtmlAgility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter.Internet
{
    public class InternetBrowser
    {
        private Parametres intParametres = null;
        private Browser intBrowser = null;
        private HtmlDocument intLastHtml = null;

        public InternetBrowser(Parametres parametres)
        {
            intParametres = parametres;
        }

        public List<Ticket> TelechargerAllTickets()
        {
            List<Ticket> tickets = new List<Ticket>();

            if (ConnexionSite())
            {
                do
                {
                    tickets.AddRange(TelechargerTicketsPage());
                } while (PageSuivante());
            }
            else
            {
                throw new FireException("Connexion au site internet en erreur");
            }

            return tickets;
        }

        public bool ConnexionSite()
        {
            string url = intParametres.AdresseMoneWeb;
            string login = intParametres.Login;
            string password = intParametres.MotDePasse;

            string formLogin = "login$ctl00$tbLogin";
            string formPassword = "login$ctl00$tbPassword";
            string eventTarget = "__EVENTTARGET";
            string eventTargetValue = "login$ctl00$btnConnexion";

            intBrowser = new Browser(url);

            intLastHtml = intBrowser.Get();
            if (intBrowser.FormElements.ContainsKey(eventTargetValue))
            {
                intBrowser.FormElements[eventTarget] = eventTargetValue;
                intBrowser.FormElements[formLogin] = login;
                intBrowser.FormElements[formPassword] = password;
                intLastHtml = intBrowser.Post();

                if (intBrowser.FormElements.ContainsKey(eventTargetValue) == false)
                {
                    return true;
                }
                else
                {
                    throw new FireException("Login / Password en erreur.");
                }
            }
            else
            {
                throw new FireException("Bouton Connexion introuvable. Erreur de site ?");
            }
        }

        public List<Ticket> TelechargerTicketsPage()
        {
            // TODO Rename
            var coco = intLastHtml.GetElementbyId("m_cphMain_ctl00_gvHistoTick");
            var truc1 = coco.SelectNodes("//td[@class='colDataDate']");
            var truc2 = coco.SelectNodes("//td[@class='colDataLibelle']");
            var truc3 = coco.SelectNodes("//td[@class='colDataChiffre']");

            List<Ticket> list = new List<Ticket>();
            if (truc1 != null && truc2 != null && truc3 != null && truc1.Count == truc2.Count && 4 * truc1.Count == truc3.Count)
            {
                for (int i = 0; i < truc1.Count; i++)
                {
                    int j = 0;
                    int i4 = i * 4;
                    list.Add(new Ticket()
                    {
                        Date = truc1[i].InnerText.Trim(),
                        Activite = truc2[i].InnerText.Trim(),
                        AncienSolde = truc3[i4 + j++].InnerText.Trim(),
                        Plateau = truc3[i4 + j++].InnerText.Trim(),
                        Financier = truc3[i4 + j++].InnerText.Trim(),
                        NouveauSolde = truc3[i4 + j++].InnerText.Trim()
                    });
                }
            }
            else
            {
                //TODO Error
            }

            return list;
        }

        public bool PageSuivante()
        {
            //string eventTarget = "__EVENTTARGET";
            //string eventTargetValue = "m$cphMain$ctl00$gvHistoTick$ctl18$ImgBtnNext";            

            ////if (intBrowser.FormElements.ContainsKey(eventTargetValue))
            //{
            //    intBrowser.FormElements[eventTarget] = eventTargetValue;
            //    intLastHtml = intBrowser.Post();

            //    return true;
            //}
            ////else
            //{
            //    throw new FireException("Bouton Next introuvable. Erreur de connexion ?");
            //}
            return false;
        }
    }
}
