using System;
using System.Configuration;

namespace CodeMetricCalculator
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime startProcessing = DateTime.Now;
            string projectPath = ConfigurationManager.AppSettings["ProjectToCalculate"];
            if (string.IsNullOrWhiteSpace(projectPath))
                Console.WriteLine("Por favor,defina o path do projeto a calcular as métricas");
            else
            {
                string outputFilePath = ConfigurationManager.AppSettings["OutputFilePath"];
                if (string.IsNullOrWhiteSpace(outputFilePath)) Console.WriteLine("Por favor,defina o path onde o arquivo gerado deverá ser salvo"); ;
                using (var vsHandler = new VisualStudioHandler(projectPath))
                {
                    try
                    {
                        if (vsHandler.CalculateMetrics())
                        {
                            var excelProcessId = vsHandler.ExportCodeMetrics();

                            if (excelProcessId != -1)
                            {
                                using (var excelHandler = new ExcelHandler(excelProcessId))
                                {
                                    Console.Write(
                                        "Informe o nome que deverá ser dado ao arquivo gerado. O arquivo será gerado em " +
                                        outputFilePath + "= ");
                                    string excelFile = Console.ReadLine();

                                    excelHandler.SaveResult(outputFilePath, excelFile);
                                    Console.WriteLine("Calculo de métricas de código concluído. Tempo gasto " + (DateTime.Now-startProcessing));
                                }
                            }
                            else Console.WriteLine("Ocorreu um erro ao export as métricas");
                        }
                        else Console.WriteLine("Ocorreu um erro no calculo das métricas");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }

            Console.ReadKey();
        }
    }
}
