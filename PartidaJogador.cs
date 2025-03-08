using System;
using System.Collections.Generic;
using System.Reflection;

namespace NBABet
{
    public class PartidaJogador
    {
        private string mData;
        public string auxData { get; set; }
        private string mAdversario { get; set; }
        public string auxAdversario { get; set; }
        public int Minutos { get; set; }
        private bool? mEmCasa { get; set; }
        private int? mFieldGoals { get; set; }
        public string auxFieldGoals { get; set; }
        private int? mFieldGoalsTentativas { get; set; }
        private int? mCestas3 { get; set; }
        public string auxCestas3 { get; set; }
        private int? mCestas3Tentativas { get; set; }
        private int? mLancesLivres { get; set; }
        public string auxLancesLivres { get; set; }
        private int? mLancesLivresTentativas { get; set; }
        private int? mCestas2 { get; set; }
        private int? mCestas2Tentativas { get; set; }
        public string Time { get; set; }
        public int Rebotes { get; set; }
        public int Assistencias { get; set; }
        public int Bloqueios { get; set; }
        public int Roubos { get; set; }
        public int Faltas { get; set; }
        public int InversoesPosse { get; set; }
        public int Pontos { get; set; }

        public string Data
        {
            get
            {
                if (mData == null)
                {
                    mData = Convert.ToDateTime(new DateTime(DateTime.Now.Year, Convert.ToInt32(auxData.Substring(0, auxData.IndexOf('/'))), Convert.ToInt32(auxData.Substring(auxData.IndexOf('/') + 1)))).ToString("yyyy/MM/dd");
                }

                return mData;
            }
            set
            {
                mData = value;
            }
        }

        public string Adversario
        {
            get
            {
                if (mAdversario == null)
                {
                    mAdversario = auxAdversario.Replace("vs", "").Replace("@", "");
                }

                return mAdversario;
            }
            set
            {
                mAdversario = value;
            }
        }

        public bool EmCasa
        {
            get
            {
                if (mEmCasa == null)
                {
                    mEmCasa = auxAdversario.StartsWith("vs");
                }

                return (bool)mEmCasa;
            }
            set
            {
                mEmCasa = value;
            }
        }

        public int FieldGoals
        {
            get
            {
                if (mFieldGoals == null)
                {
                    mFieldGoals = Convert.ToInt32(auxFieldGoals.Substring(0, auxFieldGoals.IndexOf('-')));
                }

                return (int)mFieldGoals;
            }
            set
            {
                mFieldGoals = value;
            }
        }

        public int FieldGoalsTentativas
        {
            get 
            { 
                if (mFieldGoalsTentativas == null)
                {
                    mFieldGoalsTentativas = Convert.ToInt32(auxFieldGoals.Substring(auxFieldGoals.IndexOf('-') + 1));
                }

                return (int)mFieldGoalsTentativas;
            }
            set 
            {
                mFieldGoalsTentativas = value;
            }
        }

        public int Cestas3
        {
            get 
            { 
                if (mCestas3 == null)
                {
                    mCestas3 = Convert.ToInt32(auxCestas3.Substring(0, auxCestas3.IndexOf('-')));
                }

                return (int)mCestas3;
            }
            set 
            {
                mCestas3 = value;
            }
        }

        public int Cestas3Tentativas
        {
            get 
            {
                if (mCestas3Tentativas == null)
                {
                    mCestas3Tentativas = Convert.ToInt32(auxCestas3.Substring(auxCestas3.IndexOf('-') + 1));
                }
                
                return (int)mCestas3Tentativas;
            }
            set 
            {
                mCestas3Tentativas = value;
            }
        }

        public int LancesLivres
        {
            get 
            {
                if (mLancesLivres == null)
                {
                    mLancesLivres = Convert.ToInt32(auxLancesLivres.Substring(0, auxLancesLivres.IndexOf('-')));
                }
                
                return (int)mLancesLivres;
            }
            set 
            {
                mLancesLivres = value;
            }
        }

        public int LancesLivresTentativas
        {
            get 
            {
                if (mLancesLivresTentativas == null)
                {
                    mLancesLivresTentativas = Convert.ToInt32(auxLancesLivres.Substring(auxLancesLivres.IndexOf('-') + 1));
                }
                
                return (int)mLancesLivresTentativas;
            }
            set 
            {
                mLancesLivresTentativas = value;
            }
        }

        public int Cestas2
        {
            get 
            {
                if (mCestas2 == null)
                {
                    mCestas2 = FieldGoals - Cestas3;
                }

                return (int)mCestas2; 
            }
            set 
            {
                mCestas2 = value;
            }
        }

        public int Cestas2Tentativas
        {
            get 
            {
                if (mCestas2Tentativas == null)
                {
                    mCestas2Tentativas = FieldGoalsTentativas - Cestas3Tentativas;
                }

                return (int)mCestas2Tentativas;
            }
            set 
            {
                mCestas2Tentativas = value;
            }
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
            mAdversario = (string)pAdversario;
            Minutos = (int)pMinutos;
            mEmCasa = (bool)pEmCasa;
            mFieldGoals = (int)pFieldGoals;
            mFieldGoalsTentativas = (int)pFieldGoalsTentativas;
            mCestas3 = (int)pCestas3;
            mCestas3Tentativas = (int)pCestas3Tentavias;
            mLancesLivres = (int)pLancesLivres;
            mLancesLivresTentativas = (int)pLancesLivresTentativas;
            mCestas2 = (int)pCestas2;
            mCestas2Tentativas = (int)pCestas2Tentativas;
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
