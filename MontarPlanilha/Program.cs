using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MontarPlanilha
{
    public class Program
    {
        public static void Main()
        {
            Console.WriteLine("inicio");
            if (Console.IsInputRedirected)
            {
                Console.WriteLine("entrou");
                var opIn = new List<Operacao>();
                var opOut = new List<Operacao>();
                while (true)
                {
                    var linha = Console.ReadLine();
                    Console.WriteLine($"--> {linha}");

                    if (string.IsNullOrWhiteSpace(linha)) { break; }

                    var operacao = new Operacao(linha);
                    if (operacao.Direction == "in")
                        opIn.Add(operacao);
                    else
                        opOut.Add(operacao);

                }
                //var linhas = new List<string>();

                //linhas.Add("2021.04.19 09:07:10;2;WDO$N;buy;in;1;5 576,000;2;0,00;0,00;0,00;100 000,00;C;");
                //linhas.Add("2021.04.19 09:21:14;3;WDO$N;sell;out;1;5 590,500;3;0,00;0,00;14 500,00;114 500,00;tp 5590.500;");

                //linhas.Add("2021.04.19 09:36:10;4;WDO$N;buy;in;1;5 596,500;4;0,00;0,00;0,00;114 500,00;C;");
                //linhas.Add("2021.04.19 09:42:14;5;WDO$N;sell;out;1;5 582,000;5;0,00;0,00;-14 500,00;100 000,00;sl 5582.000;");

                //foreach (var linha in linhas)
                //{
                //    if (string.IsNullOrWhiteSpace(linha)) { break; }

                //    var operacao = new Operacao(linha);
                //    if (operacao.Direction == "in")
                //        opIn.Add(operacao);
                //    else
                //        opOut.Add(operacao);
                //}
                var cout = 0;
                var tradeDia = DateTime.Now.Date;
                var nTradeDia = 0;
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("SheetOk");

                foreach (var item in opIn.OrderBy(x => x.Date).ToList())
                {
                    nTradeDia++;
                    if (item.Date.Date != tradeDia)
                    {
                        tradeDia = item.Date.Date;
                        nTradeDia = 1;
                    }

                    foreach (var iOut in opOut.Where(x => x.Date.Date == item.Date.Date).OrderBy(x => x.Date).ToList())
                    {
                        if (item.CheckOut(iOut))
                        {
                            cout++;

                            worksheet.Cell($"A{cout}").Value = item.Date.Date;
                            worksheet.Cell($"B{cout}").Value = $"{item.Date.Hour}:{item.Date.Minute}:{item.Date.Second}";
                            worksheet.Cell($"C{cout}").Value = nTradeDia;
                            worksheet.Cell($"D{cout}").Value = item.Id;
                            worksheet.Cell($"E{cout}").Value = item.Symbol;
                            worksheet.Cell($"F{cout}").Value = item.Type;
                            worksheet.Cell($"G{cout}").Value = item.Direction;
                            worksheet.Cell($"H{cout}").Value = item.Volume;
                            worksheet.Cell($"I{cout}").Value = item.Price;

                            worksheet.Cell($"J{cout}").Value = $"{iOut.Date.Hour}:{iOut.Date.Minute}:{iOut.Date.Second}";
                            worksheet.Cell($"K{cout}").Value = iOut.Id;
                            worksheet.Cell($"L{cout}").Value = iOut.Type;
                            worksheet.Cell($"M{cout}").Value = iOut.Direction;
                            worksheet.Cell($"N{cout}").Value = iOut.Price;
                            worksheet.Cell($"O{cout}").Value = iOut.Profit;
                            //worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                            //workbook.SaveAs("c:\\temp\\HelloWorld.xlsx");

                            opOut.Remove(iOut);
                            opIn.Remove(item);
                            // Console.WriteLine($"salvou planilha");
                        }
                    }
                }
                if (opIn.Count > 0 || opOut.Count > 0)
                {                    
                    var worksheetNOK = workbook.Worksheets.Add("SheetNOk");
                    cout = 0;
                    foreach (var item in opIn)
                    {
                        cout++;
                        Console.WriteLine($"Erro -->{item.Id}");
                        worksheetNOK.Cell($"A{cout}").Value = item.Date.Date;
                        worksheetNOK.Cell($"B{cout}").Value = $"{item.Date.Hour}:{item.Date.Minute}:{item.Date.Second}";
                        worksheetNOK.Cell($"C{cout}").Value = nTradeDia;
                        worksheetNOK.Cell($"D{cout}").Value = item.Id;
                        worksheetNOK.Cell($"E{cout}").Value = item.Symbol;
                        worksheetNOK.Cell($"F{cout}").Value = item.Type;
                        worksheetNOK.Cell($"G{cout}").Value = item.Direction;
                        worksheetNOK.Cell($"H{cout}").Value = item.Volume;
                        worksheetNOK.Cell($"I{cout}").Value = item.Price;
                    }
                    foreach (var item in opOut)
                    {
                        cout++;
                        Console.WriteLine($"Erro2 -->{item.Id}");
                        worksheetNOK.Cell($"A{cout}").Value = item.Date.Date;
                        worksheetNOK.Cell($"B{cout}").Value = $"{item.Date.Hour}:{item.Date.Minute}:{item.Date.Second}";
                        worksheetNOK.Cell($"C{cout}").Value = nTradeDia;
                        worksheetNOK.Cell($"D{cout}").Value = item.Id;
                        worksheetNOK.Cell($"E{cout}").Value = item.Symbol;
                        worksheetNOK.Cell($"F{cout}").Value = item.Type;
                        worksheetNOK.Cell($"G{cout}").Value = item.Direction;
                        worksheetNOK.Cell($"H{cout}").Value = item.Volume;
                        worksheetNOK.Cell($"I{cout}").Value = item.Price;
                    }
                    Console.WriteLine($"{opIn.Count+ opOut.Count} rigistros com problema! Verifiquei a SheetNOk.");
                }

                var nomeArquivo = $"PlanilhaOp-{DateTime.Now.Year}{DateTime.Now.Month}{DateTime.Now.Day}{DateTime.Now.Hour}{DateTime.Now.Minute}{DateTime.Now.Second}.xlsx";

                workbook.SaveAs(nomeArquivo);
                workbook.Dispose();
                Console.WriteLine($"Processamento concluido com sucesso! --> {nomeArquivo}");
            }
            else
            {
                Console.WriteLine("Informe um aquivo no formato csv");
            }
           
        }
    }

}
