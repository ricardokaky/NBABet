using System.Collections.Generic;

namespace NBABet
{
    public class Partida
    {
        public string DataHora { get; set; }
        public string TimeCasa { get; set; }
        public string TimeFora { get; set; }
        public string Times { get { return TimeCasa + " x " + TimeFora; } }
        public List<Jogador> Jogadores { get; set; }
        public string Url { get; set; }

        public Partida(string pDataHora, string pTimeCasa, string pTimeFora, string pUrl)
        {
            DataHora = pDataHora;
            TimeCasa = pTimeCasa;
            TimeFora = pTimeFora;
            Jogadores = new List<Jogador>();
            Url = pUrl;
        }
    }
}
