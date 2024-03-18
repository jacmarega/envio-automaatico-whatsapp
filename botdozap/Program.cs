using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Threading;
using ClosedXML.Excel;
using System.Resources;
using OpenQA.Selenium.Support.UI;
using MandarMensagem;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main()
    {
        //FERRAMENTA DE ENVIO AUTOMATICO DE MENSAGENS PARA O WHATSAPP

        string filePathName = Directory.GetCurrentDirectory() + "\\Resources\\Enviar.xlsx";
        string video = Directory.GetCurrentDirectory() + "\\Resources\\video.mp4";


        if (File.Exists(filePathName))
        {
            List<ClienteModel> listaClientes = new List<ClienteModel>();
            var workbook = new XLWorkbook(filePathName);
            // obtem apenas as linhas que foram utilizadas da planilha
            var nonEmptyDataRows = workbook.Worksheet(1).RowsUsed();

            foreach (var dataRow in nonEmptyDataRows)
            {
                if (dataRow.RowNumber() > 0)
                {
                    var lista = new ClienteModel();
                    lista.Cliente = dataRow.Cell(1).Value.ToString();
                    lista.Numero = dataRow.Cell(2).Value.ToString();
                    listaClientes.Add(lista);
                }
            }


            var driver = new ChromeDriver();

            driver.Navigate().GoToUrl("https://web.whatsapp.com/");
            Console.WriteLine("Aperte enter");
            Console.ReadLine();
            Console.WriteLine("Continuando");

            IWebElement secondMessageBox = driver.FindElement(By.XPath("(//div[@contenteditable='true'])"));
            for (int i = 0; i <= 5; i++)
            {
                Console.SetCursorPosition(5, 5);
                Thread.Sleep(1000);
                Console.WriteLine($"{i} segundos...");
            }

            foreach (var item in listaClientes)
            {

                var pessoa = item.Cliente;
                Console.WriteLine($"{pessoa}");
                var numero = item.Numero;
                var numeroFormatado = numero.Replace(" ", "").Replace("-", "");
                Console.WriteLine($"{numeroFormatado}");

                //var texto = Uri.EscapeDataString($" {pessoa} {mensagem}");
                var mensagem = "TESTANDO UMA API";
                var link = $"https://web.whatsapp.com/send?phone={numero}";
                driver.Navigate().GoToUrl(link);
                IWebElement secondMessageBo = null;
                int attempts = 0;

                while (secondMessageBo == null && attempts < 10) // tente até 10 vezes
                {
                    try
                    {
                        secondMessageBo = driver.FindElement(By.CssSelector("[title='Digite uma mensagem']"));
                        
                    }
                    catch (NoSuchElementException)
                    {
                        Console.WriteLine("Elemento não encontrado, esperando 20 segundos antes de tentar novamente");
                        for (int i = 0; i <= 20; i++)
                        {
                            Console.SetCursorPosition(5, 5);
                            Thread.Sleep(1000);
                            Console.WriteLine($"{i} segundos...");
                        }
                        attempts++;
                    }
                }

                if (secondMessageBo != null)
                {
                    IWebElement inputImage = driver.FindElement(By.CssSelector("input[type='file']"));
                    inputImage.SendKeys(video);
                    Thread.Sleep(15000);
                    IWebElement sendButton = driver.FindElement(By.CssSelector(".p357zi0d.gndfcl4n.ac2vgrno.mh8l8k0y.k45dudtp.i5tg98hk.f9ovudaz.przvwfww.gx1rr48f.f8jlpxt4.hnx8ox4h.k17s6i4e.ofejerhi.os0tgls2.g9p5wyxn.i0tg5vk9.aoogvgrq.o2zu3hjb.hftcxtij.rtx6r8la.e3b81npk.oa9ii99z.p1ii4mzz"));
                    Thread.Sleep(1000);
                    sendButton.Click();
                    secondMessageBo.Click();
                    secondMessageBo.SendKeys(mensagem);
                    secondMessageBo.Click();
                    secondMessageBo.SendKeys(Keys.Enter);
                    Thread.Sleep(20000);
                }

                else
                {
                    Console.WriteLine("Elemento não encontrado após 10 tentativas");
                }
                Console.WriteLine($"Enviado: {attempts} / {nonEmptyDataRows}");
            }
            driver.Quit();
        }
    }
}