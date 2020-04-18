using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Encog.Engine.Network.Activation;
using Encog.ML;
using Encog.ML.Data;
using Encog.ML.Data.Basic;
using Encog.Neural;
using Encog.Neural.Networks;
using Encog.Neural.Networks.Layers;
using Encog.Neural.Networks.Training;
using Encog.Neural.Networks.Training.Propagation;
using Encog.Neural.Networks.Training.Propagation.Back;

namespace ImportTXT_for_XLS
{
    class Program
    {
        #region dataset
        public static double[][] entrada =
       {
            new[] {0.492472},
            new[] {0.438347},
            new[] {0.476054},
            new[] {0.449389}
        };

        public static double[][] saida =
            {
            new[] {0.492472},
            new[] {0.438347},
            new[] {0.476054},
            new[] {0.449389}
        };
        #endregion

        static void Main(string[] args)
        {
            #region Menu
            Console.WriteLine("********* MENU - Projeto de pesquisa Redes Neurais *********");
            Console.WriteLine("1 - Importar aquivo TXT para Excel - Arquivo Casa da Borracha.");
            Console.WriteLine("2 - Importar arquivo para treinamento de Rede Neural.");
            String opcao = Console.ReadLine();
            #endregion
            #region Opcao 1
            if (opcao == "1")
            {
                #region ImportarArquivoTXTExcelCasadaBorracha
                // Inicia o componente Excel
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                //Cria uma planilha temporária na memória do computador
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                Console.WriteLine("Digite o caminho dos arquivos TXT.:");
                String x = Console.ReadLine();
                Console.WriteLine(x);
                //System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(@"C:\import");
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(x);
                int contalinha = 1;
                String linha1;
                String linha;
                int contador = 1;
                int contarquivo = 1;
                // Cria o StopWatch
                Stopwatch sw = new Stopwatch();
                foreach (System.IO.FileInfo f in di.GetFiles())
                {
                    if (f.Extension.ToLower() == ".txt")
                    {
                        //string arquivo = @"C:\import\" + f.Name;
                        string arquivo = x + f.Name;
                        if (File.Exists(arquivo))
                        {
                            try
                            {
                                using (StreamReader sr = new StreamReader(arquivo))
                                {
                                    // Lê linha por linha
                                    while ((linha1 = sr.ReadLine()) != null)
                                    {
                                        // Começa a contar o tempo
                                        sw.Start();
                                        // *** Executa a sua rotina ***
                                        linha = (linha1.Replace("\t", String.Empty)).TrimStart();
                                        if (linha.StartsWith("000"))
                                        {
                                            try
                                            {
                                                contador++; //Console.WriteLine(linha);
                                                Console.SetCursorPosition(0, Console.CursorTop);
                                                Console.Write(contador);
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.Write(linha);
                                                Console.Write(ex.Message);
                                            }

                                            var pedido = linha.Substring(0, 11);
                                            var codprod = linha.Substring(11, 15);
                                            var produto = linha.Substring(27, 50);
                                            var quantidade = linha.Substring(78, 10);
                                            string valor = "";
                                            try
                                            {
                                                valor = linha.Substring(88, 11);
                                            }
                                            catch (Exception ex)
                                            {
                                                try
                                                {
                                                    valor = linha.Substring(87, 11);
                                                }
                                                catch (Exception ex1)
                                                {
                                                    try
                                                    {
                                                        valor = linha.Substring(86, 11);
                                                    }
                                                    catch (Exception ex2)
                                                    {
                                                        try
                                                        {
                                                            valor = linha.Substring(85, 11);
                                                        }
                                                        catch (Exception ex3)
                                                        {
                                                            try
                                                            {
                                                                valor = linha.Substring(84, 11);
                                                            }
                                                            catch (Exception ex4)
                                                            {
                                                                try
                                                                {
                                                                    quantidade = linha.Substring(77, 10);
                                                                    valor = linha.Substring(83, 11);
                                                                }
                                                                catch (Exception ex5)
                                                                {
                                                                    try
                                                                    {
                                                                        quantidade = linha.Substring(77, 10);
                                                                        valor = linha.Substring(82, 11);
                                                                    }
                                                                    catch (Exception ex6)
                                                                    {
                                                                        try
                                                                        {
                                                                            quantidade = linha.Substring(75, 10);
                                                                            valor = linha.Substring(80, 11);
                                                                        }
                                                                        catch (Exception ex7)
                                                                        {
                                                                            try
                                                                            {
                                                                                quantidade = linha.Substring(74, 10);
                                                                                valor = linha.Substring(79, 11);
                                                                            }
                                                                            catch (Exception ex8)
                                                                            {
                                                                                quantidade = linha.Substring(79, 9);
                                                                                valor = linha.Substring(85, 11);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            xlWorkSheet.Cells[contalinha, 1] = f.Name.Trim();
                                            xlWorkSheet.Cells[contalinha, 2] = pedido.Trim();
                                            xlWorkSheet.Cells[contalinha, 3] = codprod.Trim();
                                            xlWorkSheet.Cells[contalinha, 4] = produto.Trim();
                                            try
                                            {
                                                xlWorkSheet.Cells[contalinha, 5] = float.Parse(quantidade.Trim());
                                                xlWorkSheet.Cells[contalinha, 6] = float.Parse(valor.Trim());
                                            }
                                            catch (Exception ex9)
                                            {
                                                xlWorkSheet.Cells[contalinha, 5] = quantidade.Trim();
                                                xlWorkSheet.Cells[contalinha, 6] = valor.Trim();
                                                Console.WriteLine("");
                                                Console.WriteLine("Erro verificar.:{0},{1},{2},{3},{4},{5}", f.Name, pedido, codprod, produto, quantidade, valor);
                                                Console.WriteLine("");
                                            }
                                            contalinha++;
                                            if (contalinha == 65001)
                                            {
                                                xlWorkBook.SaveAs("arquivo" + contarquivo + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                                                Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                                                xlWorkBook.Close(true, misValue, misValue);
                                                contarquivo++;
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.Write(ex.InnerException);
                                Console.WriteLine(ex.Message);
                            }
                        }
                        else
                        {
                            Console.WriteLine(" O arquivo " + arquivo + "não foi localizaod !");
                        }
                    }
                }
                //Salva o arquivo de acordo com a documentação do Excel.
                Console.WriteLine("Importado.: " + contador);
                // Para de contar o tempo
                sw.Stop();
                // Obtém o tempo que a rotina demorou a executar
                TimeSpan tempo = sw.Elapsed;
                Console.WriteLine("Tempo.: " + tempo);
                xlWorkBook.SaveAs("arquivo.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                //o arquivo foi salvo na pasta Meus Documentos.
                string caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Console.ReadKey();
                #endregion
            }
            #endregion
            #region Opcao 2
            else if (opcao == "2")
            {
                String linha;
                int contador = 0;
                Console.WriteLine("Digite o caminho dos arquivos TXT.:");
                String caminho = Console.ReadLine();
                Console.WriteLine(caminho);
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(caminho);
                Stopwatch sw = new Stopwatch();
                foreach (System.IO.FileInfo f in di.GetFiles())
                {
                    if (f.Extension.ToLower() == ".txt")
                    {
                        //string arquivo = @"C:\import\" + f.Name;
                        string arquivo = caminho + f.Name;
                        if (File.Exists(arquivo))
                        {
                            try
                            {
                                using (StreamReader sr = new StreamReader(arquivo))
                                {
                                    // Lê linha por linha
                                    while ((linha = sr.ReadLine()) != null)
                                    {
                                        // Começa a contar o tempo
                                        sw.Start();
                                        contador++;
                                    }
                                    #region Rede
                                    BasicNetwork rede = new BasicNetwork();
                                    rede.AddLayer(new BasicLayer(new ActivationSigmoid(), true, 1));
                                    rede.AddLayer(new BasicLayer(new ActivationSigmoid(), true, 15));
                                    rede.AddLayer(new BasicLayer(new ActivationSigmoid(), true, 1));
                                    rede.Structure.FinalizeStructure();
                                    rede.Reset();
                                    #endregion

                                    IMLDataSet dataset = new BasicMLDataSet(entrada, saida);

                                    Backpropagation propagation = new Backpropagation(rede, dataset);
                                    //ITrain learner = new Backpropagation(rede, dataset);

                                    #region treinamento
                                    int epoch = 0;
                                    while (true)
                                    {
                                        propagation.Iteration();
                                        //learner.Iteration();
                                        Console.WriteLine("Época " + epoch.ToString() + " Erro " + propagation.Error); //learner.Error);
                                        epoch++;

                                        if (epoch > 15500 || propagation.Error < 0.001)
                                            break;
                                    };
                                    #endregion

                                    #region Teste
                                    for (double t = 1; t <= 5; t++)
                                    {
                                        double num = 1.0 / t;
                                        double[] d = new double[] { num, num };
                                        //IMLData input = new BasicMLData(d);
                                        //IMLData output = rede.Compute(input);
                                        //double[] result = new double[output.Count];
                                        //output.CopyTo(result, 0, output.Count);
                                        //Console.WriteLine("{0} + {1} = {2}", num, num, result[0]);
                                    }
                                    double[] d1 = new double[] { 0.428138};
                                    IMLData input = new BasicMLData(d1);
                                    IMLData output = rede.Compute(input);
                                    double[] result = new double[output.Count];
                                    output.CopyTo(result, 0, output.Count);
                                    Console.WriteLine("{0}  {1} = {2}", input, output, result[0]);
                                    Console.ReadKey();
                                    #endregion
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                    }
                }
            }
            #endregion
        }
    }
}
