using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter.Data
{
    /// <summary>
    /// Une ligne de Log donnant des informations sur un lancement de l'application MoneWebExporter.
    /// Regroupées sous la forme d'un tableau dans le fichier Excel de travail, cela permet d'avoir les logs de lancement de l'application.
    /// </summary>
    public class LigneLog
    {
        /// <summary>
        /// La date de lancement de l'application MoneWebExporter
        /// </summary>
        public string DateLancement { get; set; }

        /// <summary>
        /// Le nombre de "ticket repas" téléchargées sur le site Web
        /// </summary>
        public int NombreNouvellesLignes { get; set; }

        /// <summary>
        /// La date de derniere relève utilisée lors du lancement
        /// </summary>
        public string DateDerniereReleveUtilisee { get; set; }

        /// <summary>
        /// Le nombre de subventions calculées sur la base des tickets repas, la date de derniere relève et la date de lancement
        /// </summary>
        public int NombreSubventions { get; set; }

        /// <summary>
        /// Un message d'erreur
        /// Vide dans le cas où tout c'est bien passé, rempli si l'application n'est pas allé au bout
        /// </summary>
        public string MessageErreur { get; set; }

        /// <summary>
        /// La pile d'appel de l'erreur
        /// Vide dans le cas où tout c'est bien passé, rempli si l'application n'est pas allé au bout
        /// </summary>
        public string PileAppelErreur { get; set; }
    }
}
