using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using System.Reflection.Metadata;
using System.Xml.Linq;

namespace TesteCetipDIDiaria
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void Test1()
        {
            // Configurando o diretório de download
            string downloadDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Downloads");
            if (!Directory.Exists(downloadDirectory))
            {
                Directory.CreateDirectory(downloadDirectory);
            }

            // Configurando as opções do EdgeDriver
            var options = new EdgeOptions();
            options.AddUserProfilePreference("download.default_directory", downloadDirectory);
            options.AddUserProfilePreference("download.prompt_for_download", false);
            options.AddUserProfilePreference("safebrowsing.enabled", true);

            IWebDriver driver = new EdgeDriver(options);
            driver.Navigate().GoToUrl("http://estatisticas.cetip.com.br/astec/series_v05/paginas/lum_web_v05_template_informacoes_di.asp?str_Modulo=completo&int_Idioma=1&int_Titulo=6&int_NivelBD=2");
            driver.Manage().Window.Maximize();
            //preenche dia inicio
            IWebElement element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[1]/div[2]/div/div[1]/input"));
            element.Clear();
            element.SendKeys("01");
            //preenche mes inicio
            element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[1]/div[2]/div/div[2]/input"));
            element.Clear();
            element.SendKeys("04");
            //preenche ano inicio
            element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[1]/div[2]/div/div[3]/input"));
            element.Clear();
            element.SendKeys("2025");
            //flegar o campo  Média  
            element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[4]/div/input[1]"));
            element.Click();
            //flegar o campo  Minimo
            element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[4]/div/input[2]"));
            element.Click();
            //flegar o campo  Maximo
            element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[4]/div/input[3]"));
            element.Click();
            //flegar o campo  Moda
            element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[4]/div/input[4]"));
            element.Click();
            //flegar o campo  Desvio Padrão
            element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[4]/div/input[5]"));
            element.Click();
            //flegar o campo  Taxa Selic
            element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[5]/div/input"));
            element.Click();
            //clicar no botão Consultar
            element = driver.FindElement(By.XPath("//*[@id=\"col_esq\"]/div/form/div[6]/div/a[2]"));
            element.Click();
            // Aguarde o download ser concluído (opcional)
            //System.Threading.Thread.Sleep(5000);
            element = driver.FindElement(By.XPath("//*[@id=\"divContainerIframeBmf\"]/div[2]/div/div/div[1]/div[2]/p/a"));
            //fazer download a partir do elemento com o xpath //*[@id="divContainerIframeBmf"]/div[2]/div/div/div[1]/div[2]/p/a
            //element = driver.FindElement(By.XPath("//*[@id=\"divContainerIframeBmf\"]/div[2]/div/div/div[1]/div[2]/p/a"));
            //como fazer o download com selenium

            //capturar o Html do elemento

            var nome = element.GetDomProperty("outerHTML");
            //Console.WriteLine(nome);
            var urlDownload = nome.Substring(nome.IndexOf("temp") + 5, 27);
            Console.WriteLine(urlDownload);

            var urlDownloadCompleto = "http://estatisticas.cetip.com.br/astec/temp/" + urlDownload;
            Console.WriteLine(urlDownloadCompleto);

            FileDownloader.DownloadFileAsync(urlDownloadCompleto, Path.Combine(downloadDirectory, urlDownload.Substring(0,15)+".xls")).Wait();

            driver.Quit();

     
        }
    }
    public class FileDownloader
    {
        public static async Task DownloadFileAsync(string fileUrl, string destinationPath)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    Console.WriteLine("Iniciando o download...");
                    // tentar fazer o download do arquivo 3 vezes
                    for (int i = 0; i < 3; i++)
                    {
                        try
                        {
                            byte[] fileBytes = await client.GetByteArrayAsync(fileUrl);
                            Console.WriteLine($"Download concluído. Arquivo salvo em: {destinationPath}");
                            await File.WriteAllBytesAsync(destinationPath, fileBytes);
                            return;
                        }
                        catch (HttpRequestException ex)
                        {
                            Console.WriteLine($"Tentativa {i + 1} falhou: {ex.Message}");
                            if (i == 2)
                            {
                                throw;
                            }
                        }
                        await Task.Delay(2000); // Espera 2 segundos antes de tentar novamente
                    }
                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Erro ao fazer o download: " + ex.Message);
                }
            }
        }
    }
}
