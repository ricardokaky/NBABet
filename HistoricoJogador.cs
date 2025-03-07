using System.Collections.Generic;

namespace NBABet
{
    public class HistoricoJogador
    {
        public List<PartidaJogador> Partidas { get; set; }
        public List<PartidaJogador> PartidasContraAdv { get; set; }
        public List<Media> Medias { get; set; }
        public List<Media> MediasContraAdv { get; set; }

        public HistoricoJogador()
        {
            Partidas = new List<PartidaJogador>();
            Medias = new List<Media>()
            {
                new Media("Cestas3Feitas", 0, 0),
                new Media("Rebotes", 0, 0),
                new Media("Assistencias", 0, 0),
                new Media("Bloqueios", 0, 0),
                new Media("Roubos", 0, 0),
                new Media("InversoesPosse", 0, 0),
                new Media("Pontos", 0, 0),
                new Media("PontosAssistenciasRebotes", 0, 0),
                new Media("PontosAssistencias", 0, 0),
                new Media("PontosRebotes", 0, 0),
                new Media("AssistenciasRebotes", 0, 0),
                new Media("RoubosBloqueios", 0, 0),
                new Media("PontosBloqueios", 0, 0),
                new Media("PontosRebotesBloqueios", 0, 0)
            };
            PartidasContraAdv = new List<PartidaJogador>();
            MediasContraAdv = new List<Media>()
            {
                new Media("Cestas3Feitas", 0, 0),
                new Media("Rebotes", 0, 0),
                new Media("Assistencias", 0, 0),
                new Media("Bloqueios", 0, 0),
                new Media("Roubos", 0, 0),
                new Media("InversoesPosse", 0, 0),
                new Media("Pontos", 0, 0),
                new Media("PontosAssistenciasRebotes", 0, 0),
                new Media("PontosAssistencias", 0, 0),
                new Media("PontosRebotes", 0, 0),
                new Media("AssistenciasRebotes", 0, 0),
                new Media("RoubosBloqueios", 0, 0),
                new Media("PontosBloqueios", 0, 0),
                new Media("PontosRebotesBloqueios", 0, 0)
            };
        }
    }
}
