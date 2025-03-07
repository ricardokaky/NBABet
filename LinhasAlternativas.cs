using System.Collections.Generic;

namespace NBABet
{
    public class LinhasAlternativas
    {
        public string Nome { get; set; }
        public List<Linha> Opcoes { get; set; }

        public LinhasAlternativas(string pNome, List<Linha> pOpcoes)
        {
            Nome = pNome;
            Opcoes = pOpcoes;
        }
    }
}
