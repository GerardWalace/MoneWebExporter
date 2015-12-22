using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MoneWebExporter.Data
{
    /// <summary>
    /// Un ticket de Sodexo
    /// Téléchargé à partir du site internet MoneWeb et enregistré dans le fichier Excel de travail
    /// </summary>
    public class Ticket
    {
        /// <summary>
        /// Information Sodexo : Date
        /// </summary>
        public string Date { get; set; }

        /// <summary>
        /// Information Sodexo : Activité
        /// </summary>
        public string Activite { get; set; }

        /// <summary>
        /// Information Sodexo : Ancien Solde
        /// </summary>
        public string AncienSolde { get; set; }

        /// <summary>
        /// Information Sodexo : Montant Plateau
        /// </summary>
        public string Plateau { get; set; }

        /// <summary>
        /// Information Sodexo : Financier
        /// </summary>
        public string Financier { get; set; }

        /// <summary>
        /// Information Sodexo : Nouveau Solde
        /// </summary>
        public string NouveauSolde { get; set; }
    }
}
