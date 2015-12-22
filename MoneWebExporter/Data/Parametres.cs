using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter.Data
{
    /// <summary>
    /// Parametres de lancement de l'application MoneWebExporter
    /// Disponible sous la forme d'un onglet dans le fichier Excel de travail
    /// </summary>
    public class Parametres
    {
        /// <summary>
        /// Adresse du site internet moneweb.fr
        /// </summary>
        public string AdresseMoneWeb { get; set; }

        /// <summary>
        /// Login du site internet moneweb.fr
        /// </summary>
        public string Login { get; set; }

        /// <summary>
        /// Mot de passe du site internet moneweb.fr
        /// </summary>
        public string MotDePasse { get; set; }

        /// <summary>
        /// Date du dernier lancement par rapport auquel il faut compter le nombre de subventions
        /// </summary>
        public string DerniereReleve { get; set; }
    }
}
