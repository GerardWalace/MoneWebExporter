using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check args
            //TODO
            string url = args[0];
            string login = args[1];
            string password = args[2];

            // Load existing Excel file
            //TODO

            // Connect to moneweb website
            //TODO
            HtmlDocument html = null;
            Browser browser = new Browser(url);            

            // Log in
            html = browser.Get();
            browser.FormElements["login$ctl00$tbLogin"] = login;
            browser.FormElements["login$ctl00$tbPassword"] = password;
            html = browser.Post();

            // Get tickets
            //TODO
            var coco = html.GetElementbyId("m_cphMain_ctl00_gvHistoTick");
            var truc1 = coco.SelectNodes("//td[@class='colDataDate']");
            var truc2 = coco.SelectNodes("//td[@class='colDataLibelle']");
            var truc3 = coco.SelectNodes("//td[@class='colDataChiffre']");

            if (truc1 != null && truc2 != null && truc3 != null && truc1.Count == truc2.Count && 4 * truc1.Count == truc3.Count)
            {
                List<SodexoTicket> list = new List<SodexoTicket>();
                for (int i = 0; i < truc1.Count; i++)
                {
                    list.Add(new SodexoTicket()
                        {
                            Date = truc1[i].InnerText.Trim(),
                            Activite = truc2[i].InnerText.Trim(),
                            AncienSolde = truc3[4 * i + 0].InnerText.Trim(),
                            Plateau = truc3[4 * i + 1].InnerText.Trim(),
                            Financier = truc3[4 * i + 2].InnerText.Trim(),
                            NouveauSolde = truc3[4 * i + 3].InnerText.Trim()
                        });
                }
            }
            else
            {
                //TODO Error
            }

            // Update Excel File
            //TODO

            // Calculate results
            //TODO
        }
    }
}
