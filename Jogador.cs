using System.Collections.Generic;

namespace NBABet
{
    public class Jogador
    {
        public string Nome { get; set; }
        public List<Linha> Linhas { get; set; }
        public List<LinhasAlternativas> LinhasAlternativas { get; set; }
        public HistoricoJogador Historico { get; set; }
        public string Time { get; set; }
        public string UrlHistorico { get; set; }

        public Jogador(string pNome)
        {
            Nome = pNome;
            Linhas = new List<Linha>();
            LinhasAlternativas = new List<LinhasAlternativas>();
        }
    }
}
