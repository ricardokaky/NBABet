namespace NBABet
{
    public class Media
    {
        public string Nome { get; set; }
        public double Valor { get; set; }
        public double CoeficienteVariacao { get; set; }

        public Media(string pNome, double pValor, double pCoeficienteVariacao)
        {
            Nome = pNome;
            Valor = pValor;
            CoeficienteVariacao = pCoeficienteVariacao;
        }
    }
}
