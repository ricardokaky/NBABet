using HtmlAgilityPack;
using Newtonsoft.Json;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace NBABet
{
    class Program
    {
        private static readonly HtmlDocument JogadorHtml = new HtmlDocument();
        private static List<HtmlNode> JogadorStats;
        static WebDriver Browser;
        private static List<Partida> Partidas;

        static void Main(string[] args)
        {
            try
            {
                ExcluiLog();

                InstanciaDriver();

                ScrapBetano();

                ProcurarHistoricoJogadores();

                GerarPlanilha();
            }
            catch (Exception ex)
            {
                GravaLog(ex.Message);
            }
            finally
            {
                EncerraProcessos();

                Console.ReadKey();
            }
        }

        private static void ExcluiLog()
        {
            if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Arquivos", "log.txt")))
            {
                File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Arquivos", "log.txt"));
            }
        }

        /// <summary>
        /// Cria um navegador do Firefox
        /// </summary>
        public static void InstanciaDriver()
        {
            try
            {
                var options = new FirefoxOptions()
                {
                    BrowserExecutableLocation = @"C:\Program Files\Mozilla Firefox\firefox.exe"
                };

                options.AddArguments(new List<string>() {
                "--disable-gpu",
                "--disable-application-cache",
                "--disable-extensions",
                "--disable-infobars",
                "--headless", //exibe ou esconde interface
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-blink-features=AutomationControlled",
                "--ignore-certificate-errors",
                "--allow-running-insecure-content",
                $"user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36" });

                FirefoxDriverService service = FirefoxDriverService.CreateDefaultService(@"C:\Program Files\Mozilla Firefox", "geckodriver.exe");

                Browser = new FirefoxDriver(service, options);
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao criar navegador: " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Varre o site da Betano procurando pelas linhas e odds de jogadores
        /// </summary>
        public static void ScrapBetano()
        {
            try
            {
                ProcurarPartidasDisponiveis();

                if (Partidas.Count() > 0)
                {
                    foreach (var partida in Partidas)
                    {
                        Browser.Navigate().GoToUrl(partida.Url);

                        ProcuraLinhasEspeciaisJogadores(partida);

                        ProcuraLinhasAlternativasJogadores(partida);
                    }
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao fazer o scrapping da Betano: " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Procura as partidas disponíveis para apostar
        /// </summary>
        private static void ProcurarPartidasDisponiveis()
        {
            try
            {
                Browser.Navigate().GoToUrl("https://br.betano.com/sport/basquete/eua/nba/17106/?bt=0");

                AplicaCookies();

                ScrollFimPagina();

                // Busca todas as partidas disponíveis que não estejam AO VIVO
                var partidasDisponiveis = Browser.FindElements(By.XPath("//div[@data-qa='sixpack' and not(.//span[text()='AO VIVO'])]"));

                Partidas = new List<Partida>();

                // Separa as informações da div e adiciona a partida
                foreach (var partida in partidasDisponiveis)
                {
                    var times = partida.FindElements(By.XPath(".//div[@data-qa='participant']//span"));

                    var timeCasa = times[0].GetAttribute("innerText");
                    var timeFora = times[1].GetAttribute("innerText");

                    var data = partida.FindElement(By.XPath("(./preceding-sibling::div[@class='tw-flex tw-flex-row tw-items-center tw-justify-start tw-w-full tw-px-nm']//span)[1]")).GetAttribute("innerText");

                    var hora = partida.FindElement(By.XPath("(.//span)[1]")).GetAttribute("innerText");

                    var dataFinal = string.Concat(data, " ", hora);

                    var linkPartida = partida.FindElement(By.XPath("(.//a)[1]")).GetAttribute("href");

                    Partidas.Add(new Partida(dataFinal, timeCasa, timeFora, linkPartida));
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao procurar as partidas disponíveis: " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Aplica os cookies do site para evitar perguntas como maior idade ou aceitação de cookies
        /// </summary>
        private static void AplicaCookies()
        {
            try
            {
                string caminhoCookies = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Arquivos", "cookies.json");

                if (File.Exists(caminhoCookies))
                {
                    var cookies = JsonConvert.DeserializeObject<List<CookieData>>(File.ReadAllText(caminhoCookies));

                    foreach (var cookie in cookies)
                    {
                        Cookie cookieData = new Cookie(
                        cookie.Name,
                        cookie.Value,
                        cookie.Domain,
                        cookie.Path,
                        cookie.Expiry,
                        cookie.Secure,
                        cookie.IsHttpOnly,
                        cookie.SameSite
                    );

                        Browser.Manage().Cookies.AddCookie(cookieData);
                    }

                    Browser.Navigate().Refresh();
                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao aplicar os cookies: " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Rola até o fim da página para garantir que todos os objetos estejam carregados
        /// </summary>
        private static void ScrollFimPagina()
        {
            try
            {
                var lastScrollYObj = Browser.ExecuteScript("return window.scrollY;");
                int lastScrollY = Convert.ToInt32(lastScrollYObj);

                int scrollStep = 100;
                int maxAttempts = 3;
                int attemptsWithoutChange = 0;

                while (attemptsWithoutChange < maxAttempts)
                {
                    Browser.ExecuteScript($"window.scrollBy(0, {scrollStep});");

                    Thread.Sleep(10);

                    var newScrollYObj = Browser.ExecuteScript("return window.scrollY;");
                    int newScrollY = Convert.ToInt32(newScrollYObj);

                    if (newScrollY == lastScrollY)
                    {
                        attemptsWithoutChange++;
                    }
                    else
                    {
                        attemptsWithoutChange = 0;
                    }

                    lastScrollY = newScrollY;
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao tentar fazer o scroll até o fim da página: " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Acessa a página das linhas especiais de jogadores e salva as linhas e odds disponíveis.
        /// </summary>
        /// <param name="partida">Partida a ser acessada</param>
        private static void ProcuraLinhasEspeciaisJogadores(Partida partida)
        {
            try
            {
                // Acha e clica na aba Especiais de jogadores
                var abaEspeciaisJogadores = Browser.FindElements(By.XPath("//div[@data-qa='playersspecials']"));
                if (!abaEspeciaisJogadores.Any())
                {
                    return;
                }

                abaEspeciaisJogadores[0].Click();

                Thread.Sleep(TimeSpan.FromSeconds(3));

                // Acha e clica no botão que expande todas as linhas
                var expandeTodasLinhas = Browser.FindElements(By.XPath("//*[local-name()='svg' and @data-qa='expand-all']"));
                if (expandeTodasLinhas.Any())
                {
                    expandeTodasLinhas[0].Click();
                }

                // Acha todas as linhas que não sejam sobre quartos
                var linhas = Browser.FindElements(By.XPath("//div[@data-marketid and not(.//span[contains(text(), 'Quarto')])]"));

                foreach (var linha in linhas)
                {
                    var nomeLinha = linha.FindElement(By.XPath(".//span[@class='table-market-header__text']")).GetAttribute("innerText");

                    // Se tiver o botão de mostrar mais linhas, clica
                    var butMostrarTodos = linha.FindElements(By.XPath(".//button[@class='load-more']"));
                    if (butMostrarTodos.Count() > 0)
                    {
                        butMostrarTodos[0].Click();
                    }

                    // Acha todas as opções da linha
                    var opcoes = linha.FindElements(By.XPath(".//div[@class='row']"));

                    foreach (IWebElement opcao in opcoes)
                    {
                        double valorLinha = 0;

                        if (nomeLinha != "Duplo-Duplo" && nomeLinha != "Triplo-Duplo")
                        {
                            valorLinha = Convert.ToDouble(opcao.FindElement(By.XPath(".//div[@class='handicap__single-item']")).GetAttribute("innerText").Replace(".", ","));
                        }

                        var nomeJogador = opcao.FindElement(By.XPath(".//div[@class='row-title']")).GetAttribute("innerText");

                        double oddOver = 0;
                        double oddUnder = 0;

                        var auxOddOver = opcao.FindElements(By.XPath(".//div[@style='--selection-column-start: 1;']"));
                        var auxOddUnder = opcao.FindElements(By.XPath(".//div[@style='--selection-column-start: 2;']"));

                        if (auxOddOver.Count() > 0)
                        {
                            oddOver = Convert.ToDouble(auxOddOver[0].GetAttribute("innerText").Replace(".", ","));
                        }

                        if (auxOddUnder.Count() > 0)
                        {
                            oddUnder = Convert.ToDouble(auxOddUnder[0].GetAttribute("innerText").Replace(".", ","));
                        }

                        if (!partida.Jogadores.Any(x => x.Nome == nomeJogador))
                        {
                            partida.Jogadores.Add(new Jogador(nomeJogador));
                        }

                        partida.Jogadores.Find(x => x.Nome == nomeJogador).Linhas.Add(new Linha(nomeLinha, valorLinha, oddOver, oddUnder));
                    }
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao procurar linhas especiais de jogadores: " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Acessa a página das linhas alternativas de jogadores e salva as linhas e odds disponíveis.
        /// </summary>
        /// <param name="partida">Partida a ser acessada</param>
        private static void ProcuraLinhasAlternativasJogadores(Partida partida)
        {
            try
            {
                // Acha e clica na aba Linhas alternativas de jogador
                var abaAlternativasJogador = Browser.FindElements(By.XPath("//div[@data-qa='alternativeplayerprops']"));
                if (!abaAlternativasJogador.Any())
                {
                    return;
                }

                abaAlternativasJogador[0].Click();

                Thread.Sleep(TimeSpan.FromSeconds(3));

                // Acha e clica no botão que expande todas as linhas
                var expandeTodasLinhas = Browser.FindElements(By.XPath("//*[local-name()='svg' and @data-qa='expand-all']"));
                if (expandeTodasLinhas.Any())
                {
                    expandeTodasLinhas[0].Click();
                }

                // Acha todas as linhas
                var linhasJogadores = Browser.FindElements(By.XPath("//div[@data-marketid]"));

                foreach (var linha in linhasJogadores)
                {
                    var titulo = linha.FindElement(By.XPath(".//div[@class='tw-self-center']")).GetAttribute("innerText");

                    string nomeJogador = "";
                    string nomeLinha = "";

                    // Separa o nome do jogador do nome da linha
                    if (titulo.ToUpper().Contains("TOTAL DE"))
                    {
                        nomeJogador = titulo.Substring(0, titulo.ToUpper().IndexOf("TOTAL DE") - 1);
                        nomeLinha = titulo.Substring(titulo.ToUpper().IndexOf("TOTAL DE"));
                    }
                    else if (titulo.ToUpper().Contains("ARREMESSOS DE"))
                    {
                        nomeJogador = titulo.Substring(0, titulo.ToUpper().IndexOf("ARREMESSOS DE") - 1);
                        nomeLinha = titulo.Substring(titulo.ToUpper().IndexOf("ARREMESSOS DE"));
                    }
                    else if (titulo.ToUpper().Contains("ROUBOS DE"))
                    {
                        nomeJogador = titulo.Substring(0, titulo.ToUpper().IndexOf("ROUBOS DE") - 1);
                        nomeLinha = titulo.Substring(titulo.ToUpper().IndexOf("ROUBOS DE"));
                    }
                    else if (titulo.ToUpper().Contains("TOCOS"))
                    {
                        nomeJogador = titulo.Substring(0, titulo.ToUpper().IndexOf("TOCOS") - 1);
                        nomeLinha = titulo.Substring(titulo.ToUpper().IndexOf("TOCOS"));
                    }

                    if (string.IsNullOrEmpty(nomeJogador) || string.IsNullOrEmpty(nomeLinha))
                    {
                        continue;
                    }

                    // Se tiver o botão de mostrar mais linhas, clica
                    var butMostrarTodos = linha.FindElements(By.XPath(".//button[@class='load-more']"));
                    if (butMostrarTodos.Count() > 0)
                    {
                        butMostrarTodos[0].Click();
                    }

                    // Acha todas as opções da linha
                    var opcoes = linha.FindElements(By.XPath(".//div[@data-qa='event-selection']"));

                    var linhas = new List<Linha>();

                    foreach (IWebElement opcao in opcoes)
                    {
                        double valorLinha = 0;

                        // Pega o valor da linha e remove o caracter "+"
                        string auxValorLinha = opcao.FindElement(By.XPath(".//span[@class='s-name']")).GetAttribute("innerText");
                        valorLinha = Convert.ToDouble(auxValorLinha.Substring(0, auxValorLinha.Length - 1));

                        var odd = Convert.ToDouble(opcao.FindElement(By.XPath(".//span[contains(@class, 'dark:tw-text-quartary')]")).GetAttribute("innerText").Replace(".", ","));

                        if (!partida.Jogadores.Any(x => x.Nome == nomeJogador))
                        {
                            partida.Jogadores.Add(new Jogador(nomeJogador));
                        }

                        linhas.Add(new Linha(nomeLinha, Convert.ToDouble(valorLinha), odd, 0));
                    }

                    partida.Jogadores.Find(x => x.Nome == nomeJogador).LinhasAlternativas.Add(new LinhasAlternativas(nomeLinha, linhas));
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao procurar linhas alternativas de jogadores: " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Acessa o site da ESPN e busca o histórico de partidas de cada jogador disponível para apostar
        /// </summary>
        public static void ProcurarHistoricoJogadores()
        {
            try
            {
                foreach (Partida partida in Partidas)
                {
                    // Procura histórico de jogadores que ainda não procurou
                    foreach (Jogador jogador in partida.Jogadores.Where(x => x.Historico == null || x.Historico.Partidas == null || !x.Historico.Partidas.Any()))
                    {
                        if (!ProcessaHistoricoJogador(jogador))
                        {
                            throw new Exception();
                        }

                        Analisar(partida, jogador);
                    }
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao procurar histórico de jogadores: " + ex.Message);
                throw ex;
            }
        }

        public static bool ProcessaHistoricoJogador(Jogador jogador)
        {
            // Pesquisa pelo jogador no site da ESPN
            Browser.Navigate().GoToUrl("https://www.espn.com/search/_/type/players/q/" + TratarNomeJogadorHistorico(jogador.Nome));

            string href;

            // Verifica se encontrou o jogador
            try
            {
                href = Browser.FindElement(By.XPath("//a[contains(@class, 'LogoTile')][.//div[contains(@class, 'LogoTile__Meta--category') and text()='NBA']]")).GetAttribute("href");
            }
            catch (Exception)
            {
                GravaLog("[ERRO] Jogador " + jogador.Nome + " não encontrado");
                return false;
            }

            jogador.UrlHistorico = href.Insert(href.IndexOf("_/"), "gamelog/");

            if (!AcessarPaginaJogador(jogador.UrlHistorico))
            {
                GravaLog("[ERRO] Falha ao acessar a página do jogador " + jogador.Nome);
                return false;
            }

            if (!ProcurarStatsJogador())
            {
                GravaLog("[ERRO] Falha ao procurar estatísticas do jogador " + jogador.Nome);
                return false;
            }

            jogador.Time = Browser.FindElement(By.XPath("//a[@class='AnchorLink clr-black']")).GetAttribute("innerText");

            ProcessarStats(jogador);

            SalvaHistorico(jogador);

            return true;
        }

        /// <summary>
        /// Converte o nome do jogador de como é exibido no site da Betano para o site da ESPN
        /// </summary>
        /// <param name="nomeJogador">Nome do jogador a ser convertido</param>
        /// <returns>Nome convertido conforme esperado pelo site da ESPN</returns>
        public static string TratarNomeJogadorHistorico(string nomeJogador)
        {
            string nomeHistorico = BancoDados.ConsultaString("select NomeHistorico from NomesJogador where NomeBet = @NomeBet", new Dictionary<string, object> { { "@NomeBet", nomeJogador } });

            if (string.IsNullOrEmpty(nomeHistorico))
            {
                nomeHistorico = nomeJogador;
            }

            return nomeHistorico.Replace(" ", "%20").Replace("-", "%20").Trim();
        }

        public bool BateuLinha(Jogador jogador, Linha linha, bool under)
        {
            try
            {
                if (linha.Nome == "Duplo-Duplo")
                {
                    return (under && !jogador.Historico.Partidas[0].DuploDuplo) || (!under && jogador.Historico.Partidas[0].DuploDuplo);
                }
                else if (linha.Nome == "Triplo-Duplo")
                {
                    return (under && !jogador.Historico.Partidas[0].TriploDuplo) || (!under && jogador.Historico.Partidas[0].TriploDuplo);
                }
                else
                {
                    // Busca o valor da linha da última partida do jogador
                    var valorPartida = Convert.ToInt32(jogador.Historico.Partidas[0].GetType().GetProperty(PropriedadeLinha(linha.Nome)).GetValue(jogador.Historico.Partidas[0]));

                    // Verifica se a linha bateu
                    return (under && valorPartida < linha.Valor) || (!under && valorPartida >= linha.Valor);
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao verificar se a linha " + linha.Nome + " do jogador " + jogador.Nome + " bateu");
                throw ex;
            }
        }

        /// <summary>
        /// Acessa a página do jogador na ESPN
        /// </summary>
        /// <param name="url">Link da página do jogador</param>
        /// <returns>True ou False dependendo se conseguiu acessar e baixar o HTML com sucesso</returns>
        private static bool AcessarPaginaJogador(string url)
        {
            Browser.Navigate().GoToUrl(url);

            JogadorHtml.LoadHtml(Browser.PageSource);

            return JogadorHtml.DocumentNode.ChildNodes.Count > 0;
        }

        /// <summary>
        /// Procura todas as estatísticas do jogador exibidas em tela
        /// </summary>
        /// <returns>True ou False dependendo se encontrou estatísticas</returns>
        private static bool ProcurarStatsJogador()
        {
            var nodos = JogadorHtml.DocumentNode.SelectNodes("//div[@class='mb5' or @class='mb4']//tr[@class='Table__TR Table__TR--sm Table__even' or " +
                                                                                                    "@class='filled Table__TR Table__TR--sm Table__even' or " +
                                                                                                    "@class='bwb-0 Table__TR Table__TR--sm Table__even' or " +
                                                                                                    "@class='filled bwb-0 Table__TR Table__TR--sm Table__even']");
            if (nodos is null || !nodos.Any())
            {
                return false;
            }

            JogadorStats = new List<HtmlNode>(nodos.ToList());

            return JogadorStats.Any();
        }

        /// <summary>
        /// Salva todas as estatísticas do jogador encontradas no site da ESPN
        /// </summary>
        /// <param name="jogador">Jogador que deve ter suas estatísticas processadas</param>
        /// <param name="ultimaPartida">Se deve procurar apenas a última partida</param>
        private static void ProcessarStats(Jogador jogador, bool ultimaPartida = false)
        {
            try
            {
                CarregaHistorico(jogador);

                foreach (HtmlNode node in JogadorStats)
                {
                    var lstNodes = node.ChildNodes.Where(x => x.Name == "td").ToList();

                    // Ignora jogos inválidos ou comemorativos
                    if (lstNodes[0].InnerText.IndexOf(" ") < 0 || lstNodes[1].InnerText.Contains("*"))
                    {
                        continue;
                    }

                    // Remove RESULT FG%, 3P% e FT%
                    lstNodes.RemoveAt(2);
                    lstNodes.RemoveAt(4);
                    lstNodes.RemoveAt(5);
                    lstNodes.RemoveAt(6);

                    var partida = ProcessarPartida(lstNodes);

                    if (jogador.Historico.Partidas.Where(x => x.Data == partida.Data).Any())
                    {
                        break;
                    }

                    jogador.Historico.Partidas.Add(partida);
                }

                int anoAnterior = DateTime.Now.Year - 1;
                int indexVirada = -1;

                for (int i = 0; i < jogador.Historico.Partidas.Count; i++)
                {
                    var data = jogador.Historico.Partidas[i].Data;

                    if (indexVirada > -1)
                    {
                        jogador.Historico.Partidas[i].Data = data.Replace(DateTime.Now.Year.ToString(), anoAnterior.ToString());
                        continue;
                    }

                    /// Verifica á partir de quando o jogo é do ano anterior
                    if (i > 0 && Convert.ToInt32(data.Substring(5, 2)) > Convert.ToInt32(jogador.Historico.Partidas[i-1].Data.Substring(5, 2)))
                    {
                        indexVirada = i;
                    }
                }

                jogador.Historico.Partidas.OrderByDescending(x => x.Data);
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao processar estatísticas do jogador " + jogador.Nome + ": " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Identifica as estatísticas da partida do histórico do jogador
        /// </summary>
        /// <param name="nodes">Nodos do HTML a serem processados</param>
        /// <returns>Partida processada</returns>
        private static PartidaJogador ProcessarPartida(List<HtmlNode> nodes)
        {
            try
            {
                var dicStats = new Dictionary<string, string>()
                {
                    { "auxData", null },
                    { "auxAdversario", null },
                    { "Minutos", null },
                    { "auxFieldGoals", null },
                    { "auxCestas3", null },
                    { "auxLancesLivres", null },
                    { "Rebotes", null },
                    { "Assistencias", null },
                    { "Bloqueios", null },
                    { "Roubos", null },
                    { "Faltas", null },
                    { "InversoesPosse", null },
                    { "Pontos", null }
                };

                for (int i = 0; i < nodes.Count(); i++)
                {
                    string text = nodes[i].InnerText;

                    if (i == 0)
                    {
                        text = text.Substring(text.IndexOf(" ") + 1);
                    }

                    dicStats[dicStats.ElementAt(i).Key] = text;
                }

                return PartidaJogador.DictionaryDePara(dicStats);
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao processar partida: " + ex.Message);
                throw ex;
            }
        }

        /// <summary>
        /// Salva o histórico do jogador
        /// </summary>
        /// <param name="jogador">Jogador sendo analisado</param>
        private static void SalvaHistorico(Jogador jogador)
        {
            int idJogador = VerificaInsereJogador(jogador);

            var ultimoJogo = BancoDados.ConsultaString("select Data from HistoricoJogador where IdJogador = @IdJogador order by Data desc limit 1", new Dictionary<string, object> { { "@IdJogador", idJogador } });

            List<PartidaJogador> partidas;

            if (string.IsNullOrEmpty(ultimoJogo))
            {
                partidas = jogador.Historico.Partidas;
            }
            else
            {
                partidas = jogador.Historico.Partidas.Where(x => DateTime.Parse(x.Data) > DateTime.Parse(ultimoJogo)).ToList();
            }

            foreach (PartidaJogador historico in partidas)
            {
                var parametros = new Dictionary<string, object>()
                {
                    { "@idJogador", idJogador },
                    { "@Data", historico.Data },
                    { "@Adversario", historico.Adversario },
                    { "@Minutos", historico.Minutos },
                    { "@EmCasa", historico.EmCasa },
                    { "@FieldGoals", historico.FieldGoals },
                    { "@FieldGoalsTentativas", historico.FieldGoalsTentativas },
                    { "@Cestas3", historico.Cestas3 },
                    { "@Cestas3Tentativas", historico.Cestas3Tentativas },
                    { "@LancesLivres", historico.LancesLivres },
                    { "@LancesLivresTentativas", historico.LancesLivresTentativas },
                    { "@Cestas2", historico.Cestas2 },
                    { "@Cestas2Tentativas", historico.Cestas2Tentativas },
                    { "@Rebotes", historico.Rebotes },
                    { "@Assistencias", historico.Assistencias },
                    { "@Bloqueios", historico.Bloqueios },
                    { "@Roubos", historico.Roubos },
                    { "@Faltas", historico.Faltas },
                    { "@InversoesPosse", historico.InversoesPosse },
                    { "@Pontos", historico.Pontos }
                };

                BancoDados.Insere("insert into HistoricoJogador values (@idJogador, @Data, @Adversario, @Minutos, @EmCasa, @FieldGoals, " +
                    "@FieldGoalsTentativas, @Cestas3, @Cestas3Tentativas, @LancesLivres, @LancesLivresTentativas, @Cestas2, @Cestas2Tentativas, " +
                    "@Rebotes, @Assistencias, @Bloqueios, @Roubos, @Faltas, @InversoesPosse, @Pontos)", parametros);
            }
        }

        private static void CarregaHistorico(Jogador jogador)
        {
            int idJogador = VerificaInsereJogador(jogador);

            jogador.Historico = new HistoricoJogador();
            jogador.Historico.Partidas.Clear();

            jogador.Historico.Partidas.AddRange(BancoDados.ConsultaTabela("select Data, Adversario, Minutos, EmCasa, FieldGoals, FieldGoalsTentativas, Cestas3, Cestas3Tentativas, LancesLivres, LancesLivresTentativas, Cestas2, " +
                                                                            "Cestas2Tentativas, Rebotes, Assistencias, Bloqueios, Roubos, Faltas, InversoesPosse, Pontos from HistoricoJogador where IdJogador = @IdJogador",
                    x => new PartidaJogador(x.GetString(0), x.GetString(1), x.GetInt32(2), x.GetBoolean(3), x.GetInt32(4), x.GetInt32(5), x.GetInt32(6), x.GetInt32(7), x.GetInt32(8), x.GetInt32(9), x.GetInt32(10), x.GetInt32(11),
                                            x.GetInt32(12), x.GetInt32(13), x.GetInt32(14), x.GetInt32(15), x.GetInt32(16), x.GetInt32(17), x.GetInt32(18)),
                new Dictionary<string, object> { { "@IdJogador", idJogador } }));

            jogador.Historico.Partidas.OrderByDescending(x => x.Data);
        }

        /// <summary>
        /// Verifica se o jogador já existe, se não insere
        /// </summary>
        /// <param name="jogador">Jogador sendo analisado</param>
        /// <returns>Id do jogador</returns>
        private static int VerificaInsereJogador(Jogador jogador)
        {
            long? idJogador = BancoDados.ConsultaInt("select Id from Jogador where Nome = @Nome", new Dictionary<string, object> { { "@Nome", jogador.Nome } });

            if (idJogador == null)
            {
                var parametros = new Dictionary<string, object>()
                {
                    { "@Nome", jogador.Nome },
                    { "@Time", jogador.Time },
                    { "@UrlHistorico", jogador.UrlHistorico }
                };

                BancoDados.Insere($"insert into Jogador (Nome, Time, UrlHistorico) values (@Nome, @Time, @UrlHistorico)", parametros);

                idJogador = BancoDados.ConsultaInt("select Id from Jogador where Nome = @Nome", new Dictionary<string, object> { { "@Nome", jogador.Nome } });
            }

            return (int)idJogador;
        }

        /// <summary>
        /// Cruza as informações das linhas de apostas com o histórico do jogador
        /// </summary>
        /// <param name="partida">Partida a ser analisada</param>
        /// <param name="jogador">Jogador a ser analisado</param>
        private static void Analisar(Partida partida, Jogador jogador)
        {
            try
            {
                string adversario = "";

                // Identifica a sigla do adversário
                if (partida.TimeCasa == jogador.Time)
                {
                    adversario = SiglaTime(partida.TimeFora);
                }
                else
                {
                    adversario = SiglaTime(partida.TimeCasa);
                }

                var jaJogouContra = jogador.Historico.Partidas.Where(x => x.Adversario == adversario).Count() > 0;

                var EhEmCasa = partida.TimeCasa == jogador.Time;

                var jaJogouEmCasaOuFora = jogador.Historico.Partidas.Where(x => x.EmCasa == EhEmCasa).Count() > 0;

                foreach (Linha linha in jogador.Linhas)
                {
                    if (linha.Nome != "Duplo-Duplo" && linha.Nome != "Triplo-Duplo")
                    {
                        // Busca lista de valores no histórico de partidas jogadas da linha sendo avaliada
                        var lst = jogador.Historico.Partidas.Select(x => Convert.ToInt32(x.GetType().GetProperty(PropriedadeLinha(linha.Nome)).GetValue(x))).ToList();

                        DefineSequenciaLinha(lst, linha);

                        linha.MediaTemporada = lst.Average();

                        // Procura histórico de partidas contra o adversário
                        if (jaJogouContra)
                        {
                            linha.MediaAdversario = jogador.Historico.Partidas.Where(x => x.Adversario == adversario).Select(x => Convert.ToInt32(x.GetType().GetProperty(PropriedadeLinha(linha.Nome)).GetValue(x))).ToList().Average();
                        }

                        // Procura histórico de partidas em casa ou fora de casa
                        if (jaJogouEmCasaOuFora)
                        {
                            linha.MediaCasaOuFora = jogador.Historico.Partidas.Where(x => x.EmCasa == EhEmCasa).Select(x => Convert.ToInt32(x.GetType().GetProperty(PropriedadeLinha(linha.Nome)).GetValue(x))).ToList().Average();
                        }

                        // Define o percentual que a linha bateu nos últimos 5 jogos
                        if (jogador.Historico.Partidas.Count() >= 5)
                        {
                            linha.Percent5PartidasOver = ((double)lst.Take(5).Where(x => x > linha.Valor).Count() / 5) * 100;
                            linha.Percent5PartidasUnder = ((double)lst.Take(5).Where(x => x < linha.Valor).Count() / 5) * 100;
                        }

                        // Define o percentual que a linha bateu nos últimos 10 jogos
                        if (jogador.Historico.Partidas.Count() >= 10)
                        {
                            linha.Percent10PartidasOver = ((double)lst.Take(10).Where(x => x > linha.Valor).Count() / 10) * 100;
                            linha.Percent10PartidasUnder = ((double)lst.Take(10).Where(x => x < linha.Valor).Count() / 10) * 100;
                        }

                        // Define o percentual que a linha bateu na temporada
                        linha.PercentTemporadaOver = ((double)lst.Where(x => x > linha.Valor).Count() / jogador.Historico.Partidas.Count()) * 100;
                        linha.PercentTemporadaUnder = ((double)lst.Where(x => x < linha.Valor).Count() / jogador.Historico.Partidas.Count()) * 100;
                    }
                    else
                    {
                        // Busca lista de valores no histórico de partidas jogadas de Duplo-Duplo ou Triplo-Duplo
                        var lst = jogador.Historico.Partidas.Select(x => Convert.ToBoolean(x.GetType().GetProperty(PropriedadeLinha(linha.Nome)).GetValue(x))).ToList();

                        DefineSequenciaDuploETriplo(lst, linha);

                        // Define o percentual que a linha bateu nos últimos 5 jogos
                        if (jogador.Historico.Partidas.Count() >= 5)
                        {
                            linha.Percent5PartidasOver = ((double)lst.Take(5).Where(x => x).Count() / 5) * 100;
                            linha.Percent5PartidasUnder = ((double)lst.Take(5).Where(x => !x).Count() / 5) * 100;
                        }

                        // Define o percentual que a linha bateu nos últimos 10 jogos
                        if (jogador.Historico.Partidas.Count() >= 10)
                        {
                            linha.Percent10PartidasOver = ((double)lst.Take(10).Where(x => x).Count() / 10) * 100;
                            linha.Percent10PartidasUnder = ((double)lst.Take(10).Where(x => !x).Count() / 10) * 100;
                        }

                        // Define o percentual que a linha bateu na temporada
                        linha.PercentTemporadaOver = ((double)lst.Where(x => x).Count() / jogador.Historico.Partidas.Count()) * 100;
                        linha.PercentTemporadaUnder = ((double)lst.Where(x => !x).Count() / jogador.Historico.Partidas.Count()) * 100;
                    }
                }

                foreach (LinhasAlternativas linhaAlternativa in jogador.LinhasAlternativas)
                {
                    // Busca lista de valores no histórico de partidas jogadas da linha sendo avaliada
                    var lst = jogador.Historico.Partidas.Select(x => Convert.ToInt32(x.GetType().GetProperty(PropriedadeLinha(linhaAlternativa.Nome)).GetValue(x))).ToList();

                    DefineSequenciaLinhaAlternativa(lst, linhaAlternativa.Opcoes);

                    // Define a média da temporada de cada linha
                    linhaAlternativa.Opcoes.ForEach(x => x.MediaTemporada = lst.Average());

                    // Procura histórico de partidas contra o adversário e define a média de cada linha
                    if (jaJogouContra)
                    {
                        linhaAlternativa.Opcoes.ForEach(x => x.MediaAdversario = jogador.Historico.Partidas.Where(y => y.Adversario == adversario).Select(y => Convert.ToInt32(y.GetType().GetProperty(PropriedadeLinha(linhaAlternativa.Nome)).GetValue(y))).ToList().Average());
                    }

                    // Procura histórico de partidas em casa ou fora de casa e define a média de cada linha
                    if (jaJogouEmCasaOuFora)
                    {
                        linhaAlternativa.Opcoes.ForEach(x => x.MediaCasaOuFora = jogador.Historico.Partidas.Where(y => y.EmCasa == EhEmCasa).Select(y => Convert.ToInt32(y.GetType().GetProperty(PropriedadeLinha(linhaAlternativa.Nome)).GetValue(y))).ToList().Average());
                    }

                    // Define o percentual que as linhas bateram nos últimos 5 jogos
                    if (jogador.Historico.Partidas.Count() >= 5)
                    {
                        linhaAlternativa.Opcoes.ForEach(x => x.Percent5PartidasOver = ((double)lst.Take(5).Where(y => y > x.Valor).Count() / 5) * 100);
                    }

                    // Define o percentual que as linhas bateram nos últimos 10 jogos
                    if (jogador.Historico.Partidas.Count() >= 10)
                    {
                        linhaAlternativa.Opcoes.ForEach(x => x.Percent10PartidasOver = ((double)lst.Take(10).Where(y => y > x.Valor).Count() / 10) * 100);
                    }

                    // Define o percentual que as linhas bateram na temporada
                    linhaAlternativa.Opcoes.ForEach(x => x.PercentTemporadaOver = ((double)lst.Where(y => y > x.Valor).Count() / jogador.Historico.Partidas.Count()) * 100);
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao analisar a partida " + partida.Times + " e jogador " + jogador.Nome);
                throw ex;
            }
        }

        /// <summary>
        /// Converte o nome da linha de como é exibido no site da Betano para como está a propriedade na classe
        /// </summary>
        /// <param name="nomeLinha">Nome da linha a ser convertida</param>
        /// <returns>Nome da propriedade correspondente</returns>
        private static string PropriedadeLinha(string nomeLinha)
        {
            string propriedade = "";

            switch (nomeLinha.ToUpper())
            {
                case "PONTOS MAIS/MENOS":
                case "TOTAL DE PONTOS":
                    propriedade = "Pontos";
                    break;
                case "REBOTES MAIS/MENOS":
                case "TOTAL DE REBOTES":
                    propriedade = "Rebotes";
                    break;
                case "ASSISTÊNCIAS MAIS/MENOS":
                case "TOTAL DE ASSISTÊNCIAS":
                    propriedade = "Assistencias";
                    break;
                case "TOTAL ARREMESSOS DE TRÊS PONTOS MARCADOS +/-":
                case "TOTAL ARREMESSOS DE TRÊS PONTOS MARCADOS":
                case "CESTAS DE TRÊS PONTOS CONVERTIDAS MAIS/MENOS":
                case "ARREMESSOS DE TRÊS PONTOS CONVERTIDOS":
                    propriedade = "Cestas3";
                    break;
                case "TOTAL DE 2 PONTOS MARCADOS MAIS/MENOS":
                case "CESTAS DE DOIS PONTOS CONVERTIDAS MAIS/MENOS":
                    propriedade = "Cestas2";
                    break;
                case "ROUBOS MAIS/MENOS":
                case "TOTAL DE ROUBOS":
                case "ROUBOS DE BOLA":
                    propriedade = "Roubos";
                    break;
                case "BLOQUEIOS MAIS/MENOS":
                case "TOTAL DE TOCOS":
                case "TOCOS":
                    propriedade = "Bloqueios";
                    break;
                case "TURNOVER MAIS/MENOS":
                case "TOTAL DE PERDAS DE BOLA MAIS/MENOS":
                    propriedade = "InversoesPosse";
                    break;
                case "PONTOS + REBOTES + ASSISTÊNCIAS MAIS/MENOS":
                case "TOTAL DE PONTOS, REBOTES E ASSISTÊNCIAS":
                    propriedade = "PontosAssistenciasRebotes";
                    break;
                case "PONTOS + REBOTES MAIS/MENOS":
                    propriedade = "PontosRebotes";
                    break;
                case "PONTOS + ASSISTÊNCIAS MAIS/MENOS":
                    propriedade = "PontosAssistencias";
                    break;
                case "REBOTES + ASSISTÊNCIAS MAIS/MENOS":
                case "TOTAL DE REBOTES E ASSISTÊNCIAS":
                    propriedade = "AssistenciasRebotes";
                    break;
                case "PONTOS + BLOQUEIOS MAIS/MENOS":
                    propriedade = "PontosBloqueios";
                    break;
                case "PONTOS, REBOTES + BLOQUEIOS MAIS/MENOS":
                    propriedade = "PontosRebotesBloqueios";
                    break;
                case "ROUBADAS + BLOQUEIOS MAIS/MENOS":
                case "ROUBOS + BLOQUEIOS MAIS/MENOS":
                    propriedade = "RoubosBloqueios";
                    break;
                case "FALTAS COMETIDAS MAIS/MENOS":
                    propriedade = "Faltas";
                    break;
                case "TENTATIVAS DE LANÇAMENTOS DE TRÊS PONTOS MAIS/MENOS":
                case "ARREMESSOS DE TRÊS PONTOS TENTADOS MAIS/MENOS":
                case "ARREMESSO DE TRÊS PONTOS TENTADOS MAIS/MENOS":
                    propriedade = "Cestas3Tentativas";
                    break;
                case "TENTATIVAS DE LANÇAMENTOS DE DOIS PONTOS MAIS/MENOS":
                case "TENTATIVAS DE LANÇAMENTO DE DOIS PONTOS MAIS/MENOS":
                    propriedade = "Cestas2Tentativas";
                    break;
                case "FIELD GOLS MARCADOS MAIS/MENOS":
                    propriedade = "FieldGoals";
                    break;
                case "TOTAL DE TENTATIVAS DE FIELD GOLS MAIS/MENOS":
                    propriedade = "FieldGoalsTentativas";
                    break;
                case "LANCES LIVRES MARCADOS MAIS/MENOS":
                    propriedade = "LancesLivres";
                    break;
                case "TOTAL DE LANCES LIVRES MAIS/MENOS":
                    propriedade = "LancesLivresTentativas";
                    break;
                case "TEMPO DE JOGO DOS JOGADORES MAIS/MENOS":
                    propriedade = "Minutos";
                    break;
                case "DOUBLE DOUBLE":
                case "DUPLO-DUPLO":
                    propriedade = "DuploDuplo";
                    break;
                case "TRIPLE DOUBLE":
                case "TRIPLO-DUPLO":
                    propriedade = "TriploDuplo";
                    break;
                default:
                    throw new Exception("[ERRO] Linha sem ocorrência no switch de propriedades.");
            }

            return propriedade;
        }

        /// <summary>
        /// Relaciona o nome do time com a sigla exibida na ESPN
        /// </summary>
        /// <param name="time">Nome do time</param>
        /// <returns>Sigla do time</returns>
        private static string SiglaTime(string time)
        {
            string sigla = BancoDados.ConsultaString("select Sigla from SiglaTime where Nome = @Nome", new Dictionary<string, object> { { "@Nome", time.ToUpper() } });

            if (string.IsNullOrEmpty(sigla))
            {
                throw new Exception("[ERRO] Time sem ocorrência no switch de siglas.");
            }

            return sigla;
        }

        /// <summary>
        /// Define a sequência dos últimos jogos do jogador que aquela linha bateu
        /// </summary>
        /// <param name="lista">Lista de valores do histórico do jogador</param>
        /// <param name="linha">Linha sendo avaliada</param>
        private static void DefineSequenciaLinha(List<int> lista, Linha linha)
        {
            try
            {
                if (lista[0] > linha.Valor)
                {
                    if (lista.Any(x => x < linha.Valor))
                    {
                        linha.SequenciaOver = lista.IndexOf(lista.First(x => x < linha.Valor));
                    }
                    else
                    {
                        linha.SequenciaOver = lista.Count();
                    }
                }
                else
                {
                    if (lista.Any(x => x > linha.Valor))
                    {
                        linha.SequenciaUnder = lista.IndexOf(lista.First(x => x > linha.Valor));
                    }
                    else
                    {
                        linha.SequenciaUnder = lista.Count();
                    }
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao definir sequencia da linha " + linha.Nome);
                throw ex;
            }
        }

        /// <summary>
        /// Define a sequência dos últimos jogos do jogador que as linhas de Duplo-Duplo ou Triplo-Duplo bateram
        /// </summary>
        /// <param name="lista">Lista de vezes que aconteceu ou não</param>
        /// <param name="linha">Linha sendo avaliada</param>
        private static void DefineSequenciaDuploETriplo(List<bool> lista, Linha linha)
        {
            try
            {
                if (lista[0])
                {
                    if (lista.Any(x => !x))
                    {
                        linha.SequenciaOver = lista.IndexOf(lista.First(x => !x));
                    }
                    else
                    {
                        linha.SequenciaOver = lista.Count();
                    }
                }
                else
                {
                    if (lista.Any(x => x))
                    {
                        linha.SequenciaUnder = lista.IndexOf(lista.First(x => x));
                    }
                    else
                    {
                        linha.SequenciaUnder = lista.Count();
                    }
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao definir sequencia de Duplo-Duplo ou Triplo-Duplo");
                throw ex;
            }
        }

        /// <summary>
        /// Define a sequência dos últimos jogos do jogador que aquela linha bateu
        /// </summary>
        /// <param name="lista">Lista de valores do histórico do jogador</param>
        /// <param name="linha">Linha sendo avaliada</param>
        private static void DefineSequenciaLinhaAlternativa(List<int> lista, List<Linha> linhas)
        {
            for (int i = 0; i < linhas.Count(); i++)
            {
                try
                {
                    if (lista.Any(x => x < linhas[i].Valor))
                    {
                        linhas[i].SequenciaOver = lista.IndexOf(lista.First(x => x < linhas[i].Valor));
                    }
                    else
                    {
                        linhas[i].SequenciaOver = lista.Count();
                    }
                }
                catch (Exception ex)
                {
                    GravaLog("[ERRO] Falha ao definir sequencia da linha alternativa " + linhas[i].Nome);
                    throw ex;
                }
            }
        }

        /// <summary>
        /// Gera uma planilha com os dados de todas as apostas disponíveis, os dados cruzados com o histórico dos jogadores e um número de avaliação final da aposta
        /// </summary>
        private static void GerarPlanilha()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Arquivos", "Linhas Assertivas.xls")))
                {
                    File.Delete(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Arquivos", "Linhas Assertivas.xls"));
                }

                // Cria a planilha
                using (var package = new ExcelPackage(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Arquivos", "Linhas Assertivas.xls")))
                {
                    var sheet = package.Workbook.Worksheets.Add("PARTIDAS");

                    int index = 1;

                    for (int i = 0; i < Partidas.Count(); i++)
                    {
                        for (int i2 = 0; i2 < Partidas[i].Jogadores.Count(); i2++)
                        {
                            for (int i3 = 0; i3 < Partidas[i].Jogadores[i2].Linhas.Count(); i3++)
                            {
                                if ((Partidas[i].Jogadores[i2].Linhas[i3].SequenciaUnder > 0 && Partidas[i].Jogadores[i2].Linhas[i3].OddUnder == 0) ||
                                    (Partidas[i].Jogadores[i2].Linhas[i3].SequenciaOver > 0 && Partidas[i].Jogadores[i2].Linhas[i3].OddOver == 0))
                                {
                                    continue;
                                }

                                var nomeLinha = Partidas[i].Jogadores[i2].Linhas[i3].Nome;

                                double oddOver = Partidas[i].Jogadores[i2].Linhas[i3].OddOver;
                                double auxPercent5Over = -1;
                                double auxPercent10Over = -1;
                                double auxPercentTemporadaOver = Partidas[i].Jogadores[i2].Linhas[i3].PercentTemporadaOver;
                                double auxSequenciaOver = Partidas[i].Jogadores[i2].Linhas[i3].SequenciaOver;

                                double oddUnder = Partidas[i].Jogadores[i2].Linhas[i3].OddUnder;
                                double auxPercent5Under = -1;
                                double auxPercent10Under = -1;
                                double auxPercentTemporadaUnder = Partidas[i].Jogadores[i2].Linhas[i3].PercentTemporadaUnder;
                                double auxSequenciaUnder = Partidas[i].Jogadores[i2].Linhas[i3].SequenciaUnder;

                                if (Partidas[i].Jogadores[i2].Linhas[i3].Percent5PartidasOver != null)
                                {
                                    auxPercent5Over = Convert.ToDouble(Partidas[i].Jogadores[i2].Linhas[i3].Percent5PartidasOver);
                                }

                                if (Partidas[i].Jogadores[i2].Linhas[i3].Percent10PartidasOver != null)
                                {
                                    auxPercent10Over = Convert.ToDouble(Partidas[i].Jogadores[i2].Linhas[i3].Percent10PartidasOver);
                                }

                                if (Partidas[i].Jogadores[i2].Linhas[i3].Percent5PartidasUnder != null)
                                {
                                    auxPercent5Under = Convert.ToDouble(Partidas[i].Jogadores[i2].Linhas[i3].Percent5PartidasUnder);
                                }

                                if (Partidas[i].Jogadores[i2].Linhas[i3].Percent10PartidasUnder != null)
                                {
                                    auxPercent10Under = Convert.ToDouble(Partidas[i].Jogadores[i2].Linhas[i3].Percent10PartidasUnder);
                                }

                                if (nomeLinha != "Duplo-Duplo" && nomeLinha != "Triplo-Duplo")
                                {
                                    double valorLinha = Partidas[i].Jogadores[i2].Linhas[i3].Valor;
                                    double mediaTemporada = Partidas[i].Jogadores[i2].Linhas[i3].MediaTemporada;
                                    double auxMediaAdversario = -1;
                                    double auxMediaCasaOuFora = -1;
                                    string mediaAdversario = "";
                                    string mediaCasaOuFora = "";

                                    if (Partidas[i].Jogadores[i2].Linhas[i3].MediaAdversario != null)
                                    {
                                        auxMediaAdversario = Convert.ToDouble(Partidas[i].Jogadores[i2].Linhas[i3].MediaAdversario);
                                        mediaAdversario = auxMediaAdversario.ToString("0.00");
                                    }

                                    if (Partidas[i].Jogadores[i2].Linhas[i3].MediaCasaOuFora != null)
                                    {
                                        auxMediaCasaOuFora = Convert.ToDouble(Partidas[i].Jogadores[i2].Linhas[i3].MediaCasaOuFora);
                                        mediaCasaOuFora = auxMediaCasaOuFora.ToString("0.00");
                                    }

                                    // Preenche uma linha para a opção OVER
                                    sheet.Cells[$"A{index}"].Value = Partidas[i].Times;
                                    sheet.Cells[$"B{index}"].Value = Partidas[i].Jogadores[i2].Nome;
                                    sheet.Cells[$"C{index}"].Value = nomeLinha;
                                    sheet.Cells[$"D{index}"].Value = valorLinha.ToString().Replace(".", ",");
                                    sheet.Cells[$"F{index}"].Value = oddOver.ToString().Replace(".", ",");
                                    sheet.Cells[$"G{index}"].Value = mediaTemporada.ToString("0.00");
                                    sheet.Cells[$"H{index}"].Value = mediaAdversario;
                                    sheet.Cells[$"I{index}"].Value = mediaCasaOuFora;
                                    sheet.Cells[$"J{index}"].Value = auxPercent5Over.ToString() + "%";
                                    sheet.Cells[$"K{index}"].Value = auxPercent10Over.ToString() + "%";
                                    sheet.Cells[$"L{index}"].Value = auxPercentTemporadaOver.ToString("0.00") + "%";
                                    sheet.Cells[$"M{index}"].Value = auxSequenciaOver;
                                    sheet.Cells[$"N{index}"].Value = AvaliacaoAposta(Partidas[i].Jogadores[i2].Linhas[i3], true).ToString("0.00");

                                    index++;

                                    // Preenche uma linha para a opção UNDER
                                    sheet.Cells[$"A{index}"].Value = Partidas[i].Times;
                                    sheet.Cells[$"B{index}"].Value = Partidas[i].Jogadores[i2].Nome;
                                    sheet.Cells[$"C{index}"].Value = nomeLinha;
                                    sheet.Cells[$"D{index}"].Value = valorLinha.ToString().Replace(".", ",");
                                    sheet.Cells[$"E{index}"].Value = "SIM";
                                    sheet.Cells[$"F{index}"].Value = oddUnder.ToString().Replace(".", ",");
                                    sheet.Cells[$"G{index}"].Value = mediaTemporada.ToString("0.00");
                                    sheet.Cells[$"H{index}"].Value = mediaAdversario;
                                    sheet.Cells[$"I{index}"].Value = mediaCasaOuFora;
                                    sheet.Cells[$"J{index}"].Value = auxPercent5Under.ToString() + "%";
                                    sheet.Cells[$"K{index}"].Value = auxPercent10Under.ToString() + "%";
                                    sheet.Cells[$"L{index}"].Value = auxPercentTemporadaUnder.ToString("0.00") + "%";
                                    sheet.Cells[$"M{index}"].Value = auxSequenciaUnder;
                                    sheet.Cells[$"N{index}"].Value = AvaliacaoAposta(Partidas[i].Jogadores[i2].Linhas[i3], false).ToString("0.00");
                                }
                                else
                                {
                                    // Preenche uma linha para a opção OVER
                                    sheet.Cells[$"A{index}"].Value = Partidas[i].Times;
                                    sheet.Cells[$"B{index}"].Value = Partidas[i].Jogadores[i2].Nome;
                                    sheet.Cells[$"C{index}"].Value = nomeLinha;
                                    sheet.Cells[$"F{index}"].Value = oddOver.ToString().Replace(".", ",");
                                    sheet.Cells[$"J{index}"].Value = auxPercent5Over.ToString() + "%";
                                    sheet.Cells[$"K{index}"].Value = auxPercent10Over.ToString() + "%";
                                    sheet.Cells[$"L{index}"].Value = auxPercentTemporadaOver.ToString("0.00") + "%";
                                    sheet.Cells[$"M{index}"].Value = auxSequenciaOver;
                                    sheet.Cells[$"N{index}"].Value = "";

                                    index++;

                                    // Preenche uma linha para a opção UNDER
                                    sheet.Cells[$"A{index}"].Value = Partidas[i].Times;
                                    sheet.Cells[$"B{index}"].Value = Partidas[i].Jogadores[i2].Nome;
                                    sheet.Cells[$"C{index}"].Value = nomeLinha;
                                    sheet.Cells[$"E{index}"].Value = "SIM";
                                    sheet.Cells[$"F{index}"].Value = oddUnder.ToString().Replace(".", ",");
                                    sheet.Cells[$"J{index}"].Value = auxPercent5Under.ToString() + "%";
                                    sheet.Cells[$"K{index}"].Value = auxPercent10Under.ToString() + "%";
                                    sheet.Cells[$"L{index}"].Value = auxPercentTemporadaUnder.ToString("0.00") + "%";
                                    sheet.Cells[$"M{index}"].Value = auxSequenciaUnder;
                                    sheet.Cells[$"N{index}"].Value = "";
                                }

                                index++;
                            }

                            for (int i3 = 0; i3 < Partidas[i].Jogadores[i2].LinhasAlternativas.Count(); i3++)
                            {
                                var opcoes = Partidas[i].Jogadores[i2].LinhasAlternativas[i3].Opcoes.ToList();

                                for (int i4 = 0; i4 < opcoes.Count(); i4++)
                                {
                                    var linha = Partidas[i].Jogadores[i2].LinhasAlternativas[i3].Opcoes.Find(x => x.Nome == opcoes[i4].Nome && x.Valor == opcoes[i4].Valor);

                                    double auxMediaAdversario = -1;
                                    double auxMediaCasaOuFora = -1;
                                    double auxPercent5Over = -1;
                                    double auxPercent10Over = -1;
                                    double auxPercentTemporadaOver = linha.PercentTemporadaOver;
                                    double auxSequenciaOver = linha.SequenciaOver;

                                    // Preenche uma linha para a opção
                                    sheet.Cells[$"A{index}"].Value = Partidas[i].Times;
                                    sheet.Cells[$"B{index}"].Value = Partidas[i].Jogadores[i2].Nome;
                                    sheet.Cells[$"C{index}"].Value = Partidas[i].Jogadores[i2].LinhasAlternativas[i3].Nome;
                                    sheet.Cells[$"D{index}"].Value = linha.Valor.ToString().Replace(".", ",");
                                    sheet.Cells[$"F{index}"].Value = linha.OddOver.ToString().Replace(".", ",");
                                    sheet.Cells[$"G{index}"].Value = linha.MediaTemporada.ToString("0.00");

                                    if (linha.MediaAdversario != null)
                                    {
                                        auxMediaAdversario = Convert.ToDouble(linha.MediaAdversario);
                                        sheet.Cells[$"H{index}"].Value = auxMediaAdversario.ToString("0.00");
                                    }

                                    if (linha.MediaCasaOuFora != null)
                                    {
                                        auxMediaCasaOuFora = Convert.ToDouble(linha.MediaCasaOuFora);
                                        sheet.Cells[$"I{index}"].Value = auxMediaCasaOuFora.ToString("0.00");
                                    }

                                    if (linha.Percent5PartidasOver != null)
                                    {
                                        auxPercent5Over = Convert.ToDouble(linha.Percent5PartidasOver);
                                    }

                                    if (linha.Percent10PartidasOver != null)
                                    {
                                        auxPercent10Over = Convert.ToDouble(linha.Percent10PartidasOver);
                                    }

                                    sheet.Cells[$"J{index}"].Value = auxPercent5Over.ToString() + "%";
                                    sheet.Cells[$"K{index}"].Value = auxPercent10Over.ToString() + "%";
                                    sheet.Cells[$"L{index}"].Value = linha.PercentTemporadaOver.ToString("0.00") + "%";
                                    sheet.Cells[$"M{index}"].Value = linha.SequenciaOver;
                                    sheet.Cells[$"N{index}"].Value = AvaliacaoAposta(linha, true).ToString("0.00");

                                    index++;
                                }
                            }
                        }
                    }

                    package.Save();
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao gerar planilha de apostas");
                throw ex;
            }
        }

        /// <summary>
        /// Define um valor de avaliação para a linha
        /// </summary>
        /// <param name="linha">Linha a ser avaliada</param>
        /// <param name="over">Se é over ou under</param>
        /// <returns>Valor de avaliação</returns>
        public static double AvaliacaoAposta(Linha linha, bool over)
        {
            double maiorMedia;
            double percentDif;
            bool mediaAFavor;

            double valorLinha = linha.Valor;

            // Verifica se é uma linha com número decimal e arredonda dependendo se for OVER ou UNDER
            if (linha.Valor % 1 != 0)
            {
                valorLinha = over ? linha.Valor + 0.5 : linha.Valor - 0.05;
            }

            maiorMedia = Math.Max(linha.MediaTemporada, valorLinha);

            // Percentual de diferença entre a média da temporada e a linha
            percentDif = Math.Abs((linha.MediaTemporada - valorLinha) / maiorMedia * 100);

            // Verifica se a média da temporada está favorável em relação à linha
            mediaAFavor = over ? (linha.MediaTemporada >= valorLinha) : (linha.MediaTemporada <= valorLinha);

            double notaMediaTemporada = mediaAFavor ? percentDif * 0.25 : percentDif * -0.25;

            double? notaMediaCasaFora = 0;

            if (linha.MediaCasaOuFora != null)
            {
                maiorMedia = Math.Max((double)linha.MediaCasaOuFora, valorLinha);

                // Percentual de diferença entre a média de casa ou fora e a linha
                percentDif = Math.Abs(((double)linha.MediaCasaOuFora - valorLinha) / maiorMedia * 100);

                // Verifica se a média de casa ou fora está favorável em relação à linha
                mediaAFavor = over ? (linha.MediaCasaOuFora >= valorLinha) : (linha.MediaCasaOuFora <= valorLinha);

                notaMediaCasaFora = mediaAFavor ? percentDif * 0.15 : percentDif * -0.15;
            }

            double? notaMediaAdversario = 0;

            if (linha.MediaAdversario != null)
            {
                maiorMedia = Math.Max((double)linha.MediaAdversario, valorLinha);

                // Percentual de diferença entre a média contra o adversário e a linha
                percentDif = Math.Abs(((double)linha.MediaAdversario - valorLinha) / maiorMedia * 100);

                // Verifica se a média contra o adversário está favorável em relação à linha
                mediaAFavor = over ? (linha.MediaAdversario >= valorLinha) : (linha.MediaAdversario <= valorLinha);

                notaMediaAdversario = mediaAFavor ? percentDif * 0.20 : percentDif * -0.20;
            }

            double? notaPercent5 = 0;

            if (over && linha.Percent5PartidasOver != null)
            {
                notaPercent5 = linha.Percent5PartidasOver * 0.20;
            }
            else if (!over && linha.Percent5PartidasUnder != null)
            {
                notaPercent5 = linha.Percent5PartidasUnder * 0.20;
            }

            double? notaPercent10 = 0;

            if (over && linha.Percent10PartidasOver != null)
            {
                notaPercent10 = linha.Percent10PartidasOver * 0.15;
            }
            else if (!over && linha.Percent10PartidasUnder != null)
            {
                notaPercent10 = linha.Percent10PartidasUnder * 0.15;
            }

            double notaPercentTemp = over ? linha.PercentTemporadaOver * 0.10 : linha.PercentTemporadaUnder * 0.10;

            double notaSequencia = over ? linha.SequenciaOver * 0.05 : linha.SequenciaUnder * 0.05;

            // Soma todas as notas das variáveis
            return Convert.ToDouble(notaMediaTemporada + notaMediaCasaFora + notaMediaAdversario + notaPercent5 + notaPercent10 + notaPercentTemp + notaSequencia);
        }

        /// <summary>
        /// Encerra os processos do SO criados durante a execução
        /// </summary>
        private static void EncerraProcessos()
        {
            try
            {
                if (Browser != null)
                {
                    Browser.Close();
                    Browser.Dispose();
                }

                foreach (var process in Process.GetProcessesByName("Firefox"))
                {
                    process.Kill();
                }

                foreach (var process in Process.GetProcessesByName("geckodriver"))
                {
                    process.Kill();
                }
            }
            catch (Exception ex)
            {
                GravaLog("[ERRO] Falha ao encerrar processos");
                throw ex;
            }
        }

        public static void GravaLog(string mensagem)
        {
            if (!File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Arquivos", "log.txt")))
            {
                using (File.Create(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Arquivos", "log.txt"))) { };
            }

            using (StreamWriter sw = File.AppendText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Arquivos", "log.txt")))
            {
                sw.WriteLine($"{DateTime.Now}: {mensagem}");
            }
        }
    }
}
