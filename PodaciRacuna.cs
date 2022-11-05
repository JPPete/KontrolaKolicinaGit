using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KolicinaBarkodTEST
{
    internal class PodaciRacuna
    {
        private string[] NazivRacuna;

        public string BrRacuna;

        public string DatumRacuna;

        public string VremeRacuna;

        public string Poslovnica;

        public string Blagajna;


        public PodaciRacuna(string nazivRacuna)
        {
            NazivRacuna = nazivRacuna.Split("_");

            BrRacuna = NazivRacuna[1];

            DatumRacuna = NazivRacuna[2].Substring(6, 2) + "/" +
                NazivRacuna[2].Substring(4, 2) + "/" + NazivRacuna[2].Substring(0, 4);

            VremeRacuna = NazivRacuna[3].Substring(0, 2) + ":" + NazivRacuna[3].Substring(2, 2) +
                ":" + NazivRacuna[3].Substring(4, 2);

            Poslovnica = NazivRacuna[4];

            Blagajna = NazivRacuna[5].Substring(0, 4);
        }

    }
}
