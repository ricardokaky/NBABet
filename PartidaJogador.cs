using System;
using System.Collections.Generic;
using System.Reflection;

namespace NBABet
{
    public class PartidaJogador
    {
        private string mData;
        public string auxAdversario { get; set; }
        public string auxFieldGoals { get; set; }
        public string auxCestas2 { get; set; }
        public string auxCestas3 { get; set; }
        public string auxLancesLivres { get; set; }
        public string Time { get; set; }
        public string Resultado { get; set; }
        public int Minutos { get; set; }
        public int Rebotes { get; set; }
        public int Assistencias { get; set; }
        public int Bloqueios { get; set; }
        public int Roubos { get; set; }
        public int Faltas { get; set; }
        public int InversoesPosse { get; set; }
        public int Pontos { get; set; }

        public string Adversario
        {
            get { return auxAdversario.Replace("vs", "").Replace("@", ""); }
        }

        public bool EmCasa
        {
            get { return auxAdversario.StartsWith("vs"); }
            set { }
        }

        public string Data
        {
            get { return Convert.ToDateTime(new DateTime(DateTime.Now.Year, Convert.ToInt32(mData.Substring(0, mData.IndexOf('/'))), Convert.ToInt32(mData.Substring(mData.IndexOf('/') + 1)))).ToString("dd/MM/yyyy"); }
            set { mData = value; }
        }

        public int FieldGoals
        {
            get { return (Convert.ToInt32(auxFieldGoals.Substring(0, auxFieldGoals.IndexOf('-')))); }
            set { }
        }

        public int FieldGoalsTentativas
        {
            get { return (Convert.ToInt32(auxFieldGoals.Substring(auxFieldGoals.IndexOf('-') + 1))); }
            set { }
        }

        public int Cestas2
        {
            get { return (FieldGoals - Convert.ToInt32(auxCestas3.Substring(0, auxCestas3.IndexOf('-')))); }
            set { }
        }

        public int Cestas2Tentativas
        {
            get { return (FieldGoalsTentativas - Convert.ToInt32(auxCestas3.Substring(auxCestas3.IndexOf('-') + 1))); }
            set { }
        }

        public int Cestas3
        {
            get { return Convert.ToInt32(auxCestas3.Substring(0, auxCestas3.IndexOf('-'))); }
            set { }
        }

        public int Cestas3Tentativas
        {
            get { return Convert.ToInt32(auxCestas3.Substring(auxCestas3.IndexOf('-') + 1)); }
            set { }
        }

        public int LancesLivres
        {
            get { return Convert.ToInt32(auxLancesLivres.Substring(0, auxLancesLivres.IndexOf('-'))); }
            set { }
        }

        public int LancesLivresTentativas
        {
            get { return Convert.ToInt32(auxLancesLivres.Substring(auxLancesLivres.IndexOf('-') + 1)); }
            set { }
        }

        public bool DuploDuplo
        {
            get
            {
                return DuploTriploDuplo >= 2;
            }
        }

        public bool TriploDuplo
        {
            get
            {
                return DuploTriploDuplo >= 3;
            }
        }

        private int DuploTriploDuplo
        {
            get
            {
                int aux = 0;

                if (Pontos >= 10)
                {
                    aux++;
                }

                if (Assistencias >= 10)
                {
                    aux++;
                }

                if (Rebotes >= 10)
                {
                    aux++;
                }

                if (Bloqueios >= 10)
                {
                    aux++;
                }

                if (Roubos >= 10)
                {
                    aux++;
                }

                return aux;
            }
        }

        public int PontosAssistenciasRebotes
        {
            get { return Pontos + Assistencias + Rebotes; }
        }

        public int PontosAssistencias
        {
            get { return Pontos + Assistencias; }
        }

        public int PontosRebotes
        {
            get { return Pontos + Rebotes; }
        }

        public int AssistenciasRebotes
        {
            get { return Assistencias + Rebotes; }
        }

        public int RoubosBloqueios
        {
            get { return Roubos + Bloqueios; }
        }

        public int PontosBloqueios
        {
            get { return Pontos + Bloqueios; }
        }

        public int PontosRebotesBloqueios
        {
            get { return Pontos + Rebotes + Bloqueios; }
        }

        public PartidaJogador() { }

        public PartidaJogador(object pData, object pAdversario, object pMinutos, object pEmCasa, object pFieldGoals, object pFieldGoalsTentativas, object pCestas3, object pCestas3Tentavias, object pLancesLivres,
            object pLancesLivresTentativas, object pCestas2, object pCestas2Tentativas, object pRebotes, object pAssistencias, object pBloqueios, object pRoubos, object pFaltas, object pInversoesPosse, object pPontos)
        {
            mData = (string)pData;
            auxAdversario = (string)pAdversario;
            Minutos = (int)pMinutos;
            EmCasa = (bool)pEmCasa;
            FieldGoals = (int)pFieldGoals;
            FieldGoalsTentativas = (int)pFieldGoalsTentativas;
            Cestas3 = (int)pCestas3;
            Cestas3Tentativas = (int)pCestas3Tentavias;
            LancesLivres = (int)pLancesLivres;
            LancesLivresTentativas = (int)pLancesLivresTentativas;
            Cestas2 = (int)pCestas2;
            Cestas2Tentativas = (int)pCestas2Tentativas;
            Rebotes = (int)pRebotes;
            Assistencias = (int)pAssistencias;
            Bloqueios = (int)pBloqueios;
            Roubos = (int)pRoubos;
            Faltas = (int)pFaltas;
            InversoesPosse = (int)pInversoesPosse;
            Pontos = (int)pPontos;
        }

        /// <summary>
        /// Transforma um dicionário de estatísticas para uma classe estruturada da partida do jogador
        /// </summary>
        /// <param name="dic">Dicionario inicial a ser transformado</param>
        /// <returns>Partida do histórico do jogador com suas estatísticas</returns>
        public static PartidaJogador DictionaryDePara(Dictionary<string, string> dic)
        {
            var partida = new PartidaJogador();

            foreach (var key in dic.Keys)
            {
                PropertyInfo property = partida.GetType().GetProperty(key);
                property.SetValue(partida, Convert.ChangeType(dic[key], property.PropertyType), null);
            }

            return partida;
        }
    }
}
