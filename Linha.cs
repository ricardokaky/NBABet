using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NBABet
{
    public class Linha
    {
        public string Nome { get; set; }
        public double Valor { get; set; }
        public double OddOver { get; set; }
        public double OddUnder { get; set; }
        public int SequenciaOver { get; set; }
        public int SequenciaUnder { get; set; }
        public double MediaTemporada { get; set; }
        public double? MediaAdversario { get; set; }
        public double? MediaCasaOuFora { get; set; }
        public double? Percent5PartidasOver { get; set; }
        public double? Percent5PartidasUnder { get; set; }
        public double? Percent10PartidasOver { get; set; }
        public double? Percent10PartidasUnder { get; set; }
        public double PercentTemporadaOver { get; set; }
        public double PercentTemporadaUnder { get; set; }

        public Linha(string pNome, double pValor)
        {
            Nome = pNome;
            Valor = pValor;
        }

        public Linha(string pNome, double pValor, double pOddOver, double pOddUnder)
        {
            Nome = pNome;
            Valor = pValor;
            OddOver = pOddOver;
            OddUnder = pOddUnder;
        }
    }
}
