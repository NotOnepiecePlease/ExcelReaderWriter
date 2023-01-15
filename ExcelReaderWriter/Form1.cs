using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using ClosedXML.Excel;
using DevExpress.LookAndFeel;
using DevExpress.Utils.Extensions;
using DevExpress.XtraSplashScreen;
using ExcelReaderWriter.Classes;
using ExcelReaderWriter.Classes_Json;
using CsvHelper.Configuration;
using System.Globalization;
using CsvHelper;
using Color = System.Drawing.Color;
using DevExpress.XtraRichEdit.Import.Html;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using static Spire.Xls.Core.Spreadsheet.HTMLOptions;

namespace ExcelReaderWriter
{
    [SuppressMessage("ReSharper", "LocalizableElement")]
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        private List<string> ESQUEMINHA = new List<string>();
        private int EXQUEMA = 3;
        private List<Site> dadosAPIs = new List<Site>();
        private static string _diretorioCsv = @"C:\ICP\api\Planilha\csv";
        private List<ClientesCSV> listOrdenadaGrowatt;
        private List<ClientesCSV> listOrdenadaRefulog;

        public Form1()
        {
            InitializeComponent();
            UserLookAndFeel.Default.SetSkinStyle(SkinSvgPalette.WXI.OfficeWhite);
        }

        private Queue<string> updateQueue = new Queue<string>();

        private void LogUsuario(string _msg)
        {
            updateQueue.Enqueue(_msg);
            //ProcessUpdateQueue();
            if (updateQueue.Count > 0)
            {
                string newText = updateQueue.Dequeue();
                logTextBox.Invoke((MethodInvoker)delegate
                {
                    logTextBox.Text += $"{DateTime.Now:dd-MM-yyyy hh-mm-ss} - {newText}\n";
                });
            }
        }

        //private void LogUsuario(string _msg)
        //{
        //    logTextBox.Text += $"{DateTime.Now:dd-MM-yyyy hh-mm-ss} - {_msg}\n";
        //}


        private async void btnCarregarDados_Click(object sender, EventArgs e)
        {
            dadosAPIs.Clear();
            OverlayWindowOptions options = new OverlayWindowOptions(
                startupDelay: 0,
                backColor: Color.DarkGray,
                opacity: 0.4,
                fadeIn: true,
                fadeOut: true,
                imageSize: new Size(70, 70)
            );

            using (IOverlaySplashScreenHandle handle = SplashScreenManager.ShowOverlayForm(this, customPainter: new CustomOverlayPainter()))
            {
                //await BuscarDadosAPI("https://energy-search.herokuapp.com/sunnyportal");
                //await BuscarDadosAPI("https://energy-search.herokuapp.com/refulogportal");
                //await BuscarDadosAPI("https://energy-search.herokuapp.com/growattportal");

                await Task.Run(() =>
                {
                    CarregarDadosCSV();
                    IniciarAPI("https://energy-search.herokuapp.com/refulogportal",
                        "https://energy-search.herokuapp.com/growattportal");

                    // Adiciona a atualização à fila
                    //updateQueue.Enqueue(newText);

                    // Processa a fila de atualizações
                    //ProcessUpdateQueue();
                });

                //MessageBox.Show("Test");

                //BuscarDadosAPI("https://energy-search.herokuapp.com/sunnyportal");
                //BuscarDadosAPI("https://energy-search.herokuapp.com/refulogportal");

                //BuscarDadosAPI(@"C:\ICP\api\growattportal.json");
                //BuscarDadosAPI(@"C:\ICP\api\refulogportal.json");
                //IniciarAPI("https://energy-search.herokuapp.com/refulogportal",
                //    "https://energy-search.herokuapp.com/growattportal");

                //http://energy-search.herokuapp.com/solarwebportal1
                //http://energy-search.herokuapp.com/solarwebportal2
            }

            await SalvarExcelNEW();
            //BuscarDadosAPI(@"C:\ICP\api\growattportal.json");
            //BuscarDadosAPI(@"C:\ICP\api\refulogportal.json");
            //BuscarDadosAPI(@"C:\ICP\api\sunnyportal.json");
            //SalvarExcelNEW();
            Thread.Sleep(0);
        }

        private void CarregarDadosCSV()
        {
            //Thread.Sleep(3000);
            //MessageBox.Show("2");

            var files = Directory.GetFiles(_diretorioCsv, "*.csv");

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                HasHeaderRecord = true,
            };

            // List<ClientesCSV> listaLocal;
            //List<ClientesCSV> dadosKgc = new List<ClientesCSV>();
            Console.WriteLine("Buscando arquivos CSV");
            LogUsuario("Buscando arquivos CSV");
            foreach (var diretorioCadaArquivo in files)
            {
                using (var reader = new StreamReader(diretorioCadaArquivo))
                using (var csv = new CsvReader(reader, config))
                {
                    //string nomeArquivo = (reader.BaseStream as FileStream).Name;
                    var dadosArquivoCSV = csv.GetRecords<ClientesCSV>();

                    // initialize the value of filename
                    string filename = null;

                    // using the method
                    filename = System.IO.Path.GetFileName(diretorioCadaArquivo);

                    Console.WriteLine($"Arquivo: {filename.Trim().TrimEnd('.', 'c', 's', 'v')} --------");
                    LogUsuario($"Arquivo: {filename.Trim().TrimEnd('.', 'c', 's', 'v')} --------");
                    if (filename == "growat.csv")
                    {
                        listOrdenadaGrowatt = dadosArquivoCSV.OrderBy(x => x.Nome).ToList();
                    }
                    else if (filename == "refusol.csv")
                    {
                        listOrdenadaRefulog = dadosArquivoCSV.OrderBy(x => x.Nome).ToList();
                    }
                }
            }

            Console.WriteLine("\n\n\n FIM DOS ARQUIVOS CSV");
            LogUsuario($"FIM DOS ARQUIVOS CSV\n\n\n");
        }

        private string BuscarPotenciaSistema(string _site, string _nomeCliente)
        {
            if (_site == "Growatt")
            {
                return listOrdenadaGrowatt.Where(x => x.Nome == _nomeCliente).Select(x => x.Potencia).FirstOrDefault();
            }
            else if (_site == "Refulog")
            {
                return listOrdenadaRefulog.Where(x => x.Nome == _nomeCliente).Select(x => x.Potencia).FirstOrDefault();
            }

            return "N/A";
        }

        private async Task SalvarExcelNEW()
        {
            try
            {
                List<Cliente> todosClientes = new List<Cliente>();
                //Juntando todos clientes em uma lista e inserindo o site de cada um no seu cadastro
                LogUsuario("Unindo clientes...");
                dadosAPIs.ForEach(x => x.clientes.ForEach(y => todosClientes.Add(new Cliente() { Site = x.nome, Cidade = y.Cidade, Nome = y.Nome, kWh = y.kWh })));

                //tratando os nomes de cada cidade pra entrar no padrao
                LogUsuario("Tratando nome das cidades...");
                foreach (var cliente in todosClientes)
                {
                    cliente.Cidade = cliente.Cidade.ToLower().RemoverEstado().RemoverPreposicao().RemoverAcentuacao().ToUpper();
                    cliente.kWh = cliente.kWh.Trim().ToLower().TrimEnd('k', 'w', 'h').Replace('.', ',');
                }

                //ordenando os clientes por nome de cidade
                LogUsuario("Ordenando clientes por cidade...");
                List<Cliente> clientesOrdenadosPorCidade = new List<Cliente>();
                clientesOrdenadosPorCidade = todosClientes.OrderBy(x => x.Cidade).ToList();

                LogUsuario("Carregando planilha de referencia...");
                var planilhaExcel = new XLWorkbook(@"C:\ICP\api\Planilha\Planilha_Monitoramento.xlsx");

                //adicionando uma aba para cada cidade na planilha excel
                //cidadesTratadas.ForEach(x => planilhaExcel.Worksheets.Add(x.ToUpper() + $" {DateTime.Now:dd-MM}"));

                LogUsuario("Lendo as abas da planilha...");
                var abasExcel = planilhaExcel.Worksheets;

                //passando por cada aba da planilha excel
                foreach (IXLWorksheet aba in abasExcel)
                {
                    //Setando formula pra celula de media
                    LogUsuario($"Lendo aba: {aba.Name}");
                    //aba.Cell($"H3").FormulaA1 = $"=AVERAGE(F3:F{aba.RowsUsed().Count() + 1})";

                    //Serando as formatacoes condicionais, celula formatada de acordo com um calculo
                    //aba.Range($"G3:G{aba.RowsUsed().Count()}").AddConditionalFormat().WhenBetween("-0.2", "-0.1")
                    //    .Fill.SetBackgroundColor(XLColor.YellowMunsell);
                    aba.Range($"G3:G100").AddConditionalFormat().WhenBetween("-0.2", "-0.1")
                        .Fill.SetBackgroundColor(XLColor.YellowMunsell);

                    //Link sobre a formula =$G3 <- 20% : https://www.ablebits.com/office-addins-blog/relative-absolute-reference-excel/
                    //Basicamente =$G3 <-- Significa que a coluna vai ser travada, sempre sera G e o valor muda a cada linha (o valor é o numero da linha)
                    //inicia na linha 3 porque é a primeira a ter valor
                    //aba.Range($"G3:G{aba.RowsUsed().Count()}").AddConditionalFormat().WhenIsTrue("=$G3<-20%")
                    //    .Fill.SetBackgroundColor(XLColor.BrickRed); 
                    aba.Range($"G3:G100").AddConditionalFormat().WhenIsTrue("=$G3<-20%")
                        .Fill.SetBackgroundColor(XLColor.BrickRed);

                    //aba.Range($"G3:G{aba.RowsUsed().Count()}").AddConditionalFormat().WhenIsTrue("=$G3>-10%")
                    //    .Fill.SetBackgroundColor(XLColor.NoColor);
                    aba.Range($"G3:G100").AddConditionalFormat().WhenIsTrue("=$G3>-10%")
                        .Fill.SetBackgroundColor(XLColor.NoColor);

                    //Setado a aba de data
                    aba.Cell($"D1").Value = $"DIA {DateTime.Now.ToString("dd/MM")}";

                    LogUsuario($"Carregando dados da API no Excel");
                    foreach (Cliente cliente in clientesOrdenadosPorCidade)
                    {
                        if (aba.Name == cliente.Cidade + $" {DateTime.Now:dd-MM}")
                        {
                            if (!ESQUEMINHA.Contains(cliente.Cidade))
                            {
                                clientesOrdenadosPorCidade.Where(y => aba.Name == y.Cidade + $" {DateTime.Now:dd-MM}").ForEach(x =>
                                {
                                    aba.Cell($"B{EXQUEMA}").Value = x.Nome;
                                    EXQUEMA++;
                                });
                                ESQUEMINHA.Add(cliente.Cidade);
                                EXQUEMA = 3;
                                //aba.Cell($"H3").FormulaA1 = $"=AVERAGE(F3:F{aba.RowsUsed().Count() + 1})";
                            }

                            int quantidadeClientesNaCidade = clientesOrdenadosPorCidade.Count(x => x.Cidade + $" {DateTime.Now:dd-MM}" == aba.Name);
                            int linhaLidaExcel = 3;
                            for (int contadorLinha1 = 0; contadorLinha1 < quantidadeClientesNaCidade; contadorLinha1++)
                            {
                                string nomeLinha = aba.Cell($"B{linhaLidaExcel}").Value.ToString();
                                string nomeLista = cliente.Nome;
                                if (nomeLinha.ToLower() == nomeLista.ToLower())
                                {
                                    if (BuscarPotenciaSistema(cliente.Site, cliente.Nome) == null)
                                    {
                                        //aba.Range($"A{linhaLidaExcel}:H{linhaLidaExcel}").Delete(XLShiftDeletedCells.ShiftCellsUp);
                                        //aba.Row(linhaLidaExcel).Delete();
                                        aba.Range($"A{linhaLidaExcel}:G{linhaLidaExcel}").Delete(XLShiftDeletedCells.ShiftCellsUp);
                                        linhaLidaExcel++;
                                        break;
                                    }

                                    //Setando as formulas
                                    aba.Cell($"F{linhaLidaExcel}").FormulaA1 = $"=(D{linhaLidaExcel}/E{linhaLidaExcel})";
                                    aba.Cell($"F{linhaLidaExcel}").Style.NumberFormat.NumberFormatId = 2;
                                    aba.Cell($"G{linhaLidaExcel}").FormulaA1 = $"=(F{linhaLidaExcel}-$H$3)/$H$3";
                                    aba.Cell($"G{linhaLidaExcel}").Style.NumberFormat.NumberFormatId = 10;

                                    //Setando os valores
                                    aba.Cell($"A{linhaLidaExcel}").Value = cliente.Site;
                                    aba.Cell($"B{linhaLidaExcel}").Value = cliente.Nome;
                                    aba.Cell($"C{linhaLidaExcel}").Value = cliente.Cidade;
                                    aba.Cell($"D{linhaLidaExcel}").SetValue(cliente.kWh);
                                    aba.Cell($"E{linhaLidaExcel}").SetValue(BuscarPotenciaSistema(cliente.Site, cliente.Nome).Replace('.', ',') ?? "A");

                                    //tratamento para primeira linha nao por style sem dados
                                    if (!string.IsNullOrWhiteSpace(aba.Cell($"A{linhaLidaExcel}").GetValue<string>()))
                                    {
                                        //Setando as bordas e alinhamento do texto
                                        aba.Cell($"A{linhaLidaExcel}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                        aba.Cell($"B{linhaLidaExcel}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                        aba.Cell($"C{linhaLidaExcel}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                        aba.Cell($"C{linhaLidaExcel}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                        aba.Cell($"D{linhaLidaExcel}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                        aba.Cell($"E{linhaLidaExcel}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                        aba.Cell($"E{linhaLidaExcel}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                        aba.Cell($"F{linhaLidaExcel}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                        aba.Cell($"G{linhaLidaExcel}").Style.Border.RightBorder = XLBorderStyleValues.Thin;
                                    }


                                    //Setando background de acordo com o cliente
                                    if (cliente.Site == "Refulog")
                                    {
                                        aba.Cell($"A{linhaLidaExcel}").Style.Fill.SetBackgroundColor(XLColor.Redwood);
                                    }
                                    else if (cliente.Site == "Growatt")
                                    {
                                        aba.Cell($"A{linhaLidaExcel}").Style.Fill.SetBackgroundColor(XLColor.GreenPigment);
                                    }
                                    else if (cliente.Site == "Sunny")
                                    {
                                        aba.Cell($"A{linhaLidaExcel}").Style.Fill.SetBackgroundColor(XLColor.YellowMunsell);
                                    }

                                    aba.Cell($"H3").FormulaA1 = $"=AVERAGE(F3:F{aba.RowsUsed().Count() + 1})";
                                }

                                //if (quantidadeClientesNaCidade + 2 == linhaLidaExcel)
                                //{
                                //aba.Cell($"G{linhaLidaExcel}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                //}

                                linhaLidaExcel++;
                            }
                        }
                    }

                    //Setando o tracinha em baixo da ultima linha da coluna G
                    int ultimalinha = aba.LastRowUsed().RowNumber();
                    if (!string.IsNullOrWhiteSpace(aba.Cell($"A{ultimalinha}").GetValue<string>()))
                    {
                        aba.Cell($"G{ultimalinha}").Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    }
                }


                LogUsuario("Calculando formulas...");
                planilhaExcel.CalculateMode = XLCalculateMode.Auto;
                LogUsuario("Salvando planilhas...");
                planilhaExcel.SaveAs(@"C:\ICP\api\Planilha\Planilha_concluida.xlsx");
            }
            catch (Exception e)
            {
                MessageBox.Show($"{e}");
            }
        }

        private async Task BuscarDadosAPI(string _linkApi)
        {
            try
            {
                //HttpClient client = new HttpClient { BaseAddress = new Uri(_linkApi) };
                //LogUsuario($"Buscando na API: {_linkApi}");

                //HttpResponseMessage response = await client.GetAsync("");
                //string content = await response.Content.ReadAsStringAsync();
                string content = File.ReadAllText(_linkApi);

                // Escrever(_linkApi, content);
                Console.WriteLine(content);
                Console.WriteLine("\n\n\n TESTEEEEEEEEEEEEEEEEEEEE\n\n\n");
                //file.WriteLineAsync(content);
                Site siteDados = JsonConvert.DeserializeObject<Site>(content);

                dadosAPIs.Add(siteDados);
                LogUsuario($"Carregado: {siteDados.nome}");
            }
            catch (Exception e)
            {
                MessageBox.Show($"{e}");
            }
        }

        public async Task IniciarAPI(string _linkApi1, string _linkApi2)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile($"appsettings.json");

            var config = builder.Build();

            var _urlBase = config.GetSection("API_Access:UrlBase").Value;

            //https://energy-search.herokuapp.com/login
            var _uri = _urlBase + "login";
            //var _uri = "https://energy-search.herokuapp.com/login";

            var _email = config.GetSection("API_Access:email").Value;
            //var _email = "superman@gmail1.com";
            var _password = config.GetSection("API_Access:password").Value;
            //var _password = "qwer12";

            LogUsuario("Dados: uri - email e senha");
            Console.WriteLine("Dados: uri - email e senha");
            LogUsuario($"{_uri} -  {_email} - {_password} \n");
            Console.WriteLine($"{_uri} -  {_email} - {_password} \n");

            using (var client = new HttpClient())
            {
                //limpa o header
                client.DefaultRequestHeaders.Accept.Clear();

                //incluir o cabeçalho Accept que será envia na requisição
                client.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));

                // Envio da requisição a fim de autenticar
                // e obter o token de acesso
                var credenciaisJson = JsonConvert.SerializeObject(new
                {
                    login = _email,
                    password = _password
                });

                HttpResponseMessage respToken = client.PostAsync(_uri, new StringContent(
                    credenciaisJson, Encoding.UTF8, "application/json")).Result;

                //obtem o token gerado
                string conteudo = respToken.Content.ReadAsStringAsync().Result;

                Console.WriteLine(conteudo + "\n");

                if (respToken.StatusCode == HttpStatusCode.OK)
                {
                    //deserializa o token e data de expiração para o objeto Token
                    //Token token = JsonConvert.DeserializeObject<Token>(conteudo);
                    var token = JsonConvert.DeserializeObject(conteudo);

                    Console.WriteLine("A U T E N T I C A D O \n");
                    LogUsuario("A U T E N T I C A D O \n");

                    // Associar o token aos headers do objeto
                    // do tipo HttpClient
                    //client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.AccessToken);
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.ToString());

                    Console.WriteLine("Consultando a API - Energy LOGIN");
                    LogUsuario("Consultando a API - Energy LOGIN");
                    Console.WriteLine("==========================================");

                    //Aqui eu vou chamar o metodo de sinho pra pegar o retorno
                    await BuscarDadosAPI(client, _linkApi1);
                    LogUsuario("API 1: " + _linkApi1);
                    await BuscarDadosAPI(client, _linkApi2);
                    LogUsuario("API 2: " + _linkApi2);
                }
            }
        }

        private async Task BuscarDadosAPI(HttpClient client, string _linkApi)
        {
            try
            {
                HttpResponseMessage response = client.GetAsync($"{_linkApi}").Result;
                LogUsuario($"Buscando na API: {_linkApi}");

                if (response.StatusCode == HttpStatusCode.OK)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    Site siteDados = JsonConvert.DeserializeObject<Site>(content);
                    //File.WriteAllText($"{_linkApi}.json", content);

                    //using (StreamWriter writetext = new StreamWriter($"{_linkApi}.json"))
                    //{
                    //    writetext.WriteAsync(content);
                    //}

                    Console.WriteLine(content);
                    Console.WriteLine("\n\n\n TESTEEEEEEEEEEEEEEEEEEEE\n\n\n");

                    dadosAPIs.Add(siteDados);
                    LogUsuario($"Carregado: {siteDados.nome}");
                }
                else
                {
                    Console.WriteLine("Token provavelmente expirado!");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show($"{e}");
            }
        }

        private void btnSalvarDados_Click(object sender, EventArgs e)
        {
            var cidadesTratadas =
                dadosAPIs
                    .SelectMany(z => z.clientes)
                    .Select(y => y.Cidade.ToLower().RemoverEstado().RemoverPreposicao().RemoverAcentuacao())
                    .OrderBy(c => c).Distinct();

            var clientes = dadosAPIs.SelectMany(x => x.clientes);

            clientes.ForEach(y =>
            {
                File.AppendAllText($".\\Cidades\\{y.Cidade.RemoverAcentuacao().RemoverEstado().RemoverPreposicao()}.txt", y.Nome + "\n");
            });
        }
    }
}