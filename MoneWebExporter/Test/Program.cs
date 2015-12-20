using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter.Test
{
    class Program
    {
        static void OldMain(string[] args)
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
                        
            List<SodexoTicket> list = new List<SodexoTicket>();
            if (truc1 != null && truc2 != null && truc3 != null && truc1.Count == truc2.Count && 4 * truc1.Count == truc3.Count)
            {
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
            FileInfo xlFile = new FileInfo("test.xlsx");
            using (ExcelPackage xlPackage = new ExcelPackage(xlFile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.Add("Test Export");
                worksheet.Cells[1, 1].Value = "Date";
                worksheet.Cells[1, 2].Value = "Activité";
                worksheet.Cells[1, 3].Value = "Ancien solde";
                worksheet.Cells[1, 4].Value = "Plateau";
                worksheet.Cells[1, 5].Value = "Financier";
                worksheet.Cells[1, 6].Value = "Nouveau solde";

                for (int i = 0; i < list.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = list[i].Date;
                    worksheet.Cells[i + 2, 2].Value = list[i].Activite;
                    worksheet.Cells[i + 2, 3].Value = list[i].AncienSolde;
                    worksheet.Cells[i + 2, 4].Value = list[i].Plateau;
                    worksheet.Cells[i + 2, 5].Value = list[i].Financier;
                    worksheet.Cells[i + 2, 6].Value = list[i].NouveauSolde;
                }
                                    
                xlPackage.Save();
            }

            // Calculate results
            //TODO
        }
    }
}
