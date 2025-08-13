using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LettreProfessionel.Models
{
    public class LettreModel
    {
        public int Index { get; set; }
        public string? Titre { get; set; }
        public string? CheminFichier { get; set; }

        public string Affichage => $"{Index}. {Titre}";

        public override string ToString() => Affichage;
    }
}
