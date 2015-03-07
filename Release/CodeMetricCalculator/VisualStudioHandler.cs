using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using EnvDTE;
using Excel = Microsoft.Office.Interop.Excel;

namespace CodeMetricCalculator
{
    public class VisualStudioHandler : IDisposable
    {
        public void Dispose()
        {
            KillInstance();
        }

        public VisualStudioHandler(string solutionFile)
        {
            var startOpenVS = DateTime.Now;
            Console.WriteLine("Abrindo uma instância de Visual Studio");
            ActiveVsApps = GetActiveVSApps();
            this.Dte = (DTE)Activator.CreateInstance(Type.GetTypeFromProgID("VisualStudio.DTE.12.0"));
            this.SolutionFile = solutionFile;
            ExportCodeMetricsResult = -1;
            Console.WriteLine("Tempo gasto abrindo o Visual Studio: " + (DateTime.Now - startOpenVS));
        }

        /// <summary>
        /// The Visual Studio IDE instance.
        /// </summary>
        private DTE Dte { get; set; }

        private List<int> ActiveVsApps { get; set; }

        /// <summary>
        /// Indicates whether the code metrics is calculated
        /// </summary>
        public bool CodeMetricsCalculated { get; set; }

        /// <summary>
        /// The result of the ExportCodeMetrics function. When -1 the operation has failed, else it is the process id of the Excel application
        /// that contains the code metrics.
        /// </summary>
        public int ExportCodeMetricsResult { get; set; }

        /// <summary>
        /// The full path to the solution that is used to open
        /// </summary>
        public string SolutionFile { get; set; }

        /// <summary>
        /// Indicates whether the solution is opened
        /// </summary>
        public bool SolutionOpened { get; set; }

        public bool CalculateMetrics()
        {
            if (OpenSolution()) return CalculateCodeMetrics();
            return false;
        }

        private int MaxRetries
        {
            get
            {
                return 50;
            }
        }
        private UIHierarchyItem SolutionNode
        {
            get
            {
                var rootNode = (UIHierarchy)Dte.Windows.Item(DTEConstants.vsWindowKindSolutionExplorer).Object;
                return rootNode.UIHierarchyItems.Item(1);
            }
        }

        /// <summary>
        /// Opens the solution in Visual Studio.
        /// </summary>
        private bool OpenSolution()
        {
            int retries;

            try
            {

                Console.WriteLine("Abrindo solução {0}", SolutionFile);
                // Open the solution
                retries = 0;
                var startOpenSolution = DateTime.Now;
                while (true)
                {
                    try
                    {
                        Dte.Solution.Open(SolutionFile);
                        break;
                    }
                    catch (COMException)
                    {
                        System.Threading.Thread.Sleep(5000);
                        if (retries < MaxRetries)
                            retries += 1;
                        else
                            break;
                    }
                }
                var timeElapsedOpenSolution = DateTime.Now - startOpenSolution;
                Console.WriteLine("Tempo para abrir a solução: " + timeElapsedOpenSolution);

                // Wait until ready
                retries = 0;
                while (true)
                {
                    try
                    {
                        if (Dte.Solution.IsOpen)
                            break;
                    }
                    catch (COMException)
                    {
                        System.Threading.Thread.Sleep(5000);
                        if (retries < MaxRetries)
                            retries += 1;
                        else
                        {
                            Console.WriteLine("Não foi possível abrir a solução no caminho específicado. Por favor, confira o caminho informado no seu arquivo de configuração");
                            return false;
                        }
                    }
                }
            }
            catch (ThreadAbortException)
            {
                Quit();
            }

            SolutionOpened = true;
            return true;
        }

        private bool CalculateCodeMetrics()
        {
            int retries;

            CodeMetricsCalculated = false;

            // Start the build.
            Console.WriteLine("Compilando solução");
            var startBuild = DateTime.Now;
            retries = 0;
            while (true)
            {
                try
                {
                    Dte.Solution.SolutionBuild.Build(true);
                    break;
                }
                catch (COMException)
                {
                    System.Threading.Thread.Sleep(5000);
                    if (retries < MaxRetries)
                        retries += 1;
                    else
                        break;
                }
            }

            var timeElapsedToBuild = DateTime.Now - startBuild;
            Console.WriteLine("Tempo gasto compilando: " + timeElapsedToBuild);

            // When the build failed, return false
            retries = 0;
            while (true)
            {
                try
                {
                    if (Dte.Solution.SolutionBuild.LastBuildInfo != 0)
                    {
                        Console.WriteLine("Ocorreu um erro ao compilar a solução solicitada. Por favor, confira a integridade do código antes de calcular as métricas de códig.");
                        return false;
                    }
                    break;
                }
                catch (COMException)
                {
                    System.Threading.Thread.Sleep(5000);
                    if (retries < MaxRetries)
                        retries += 1;
                    else
                        Console.WriteLine("Ocorreu um erro ao compilar a solução solicitada. Por favor, confira a integridade do código antes de calcular as métricas de código.");
                    return false;
                }
            }

            // Select the solution in the "Solution Explorer"
            retries = 0;
            while (true)
            {
                try
                {
                    SolutionNode.Select(vsUISelectionType.vsUISelectionTypeSelect);
                    break;
                }
                catch (COMException)
                {
                    System.Threading.Thread.Sleep(5000);
                    if (retries < MaxRetries)
                        retries += 1;
                    else
                    {
                        Console.WriteLine("Impossível selecionar a solução no Solution Explorer");
                        return false;
                    }
                }
            }

            retries = 0;
            while (true)
            {
                try
                {
                    Dte.ExecuteCommand("ProjectandSolutionContextMenus.Solution.CalculateCodeMetrics", String.Empty);
                    break;
                }
                catch (COMException)
                {
                    System.Threading.Thread.Sleep(5000);
                    if (retries < MaxRetries)
                        retries += 1;
                    else
                    {
                        Console.WriteLine("Ocorreu um erro ao executar o calculo das métricas de código");
                        return false;
                    }
                }
            }

            CodeMetricsCalculated = true;
            return true;
        }

        public void Quit()
        {
            try
            {
                Dte.Quit();
            }
            catch (Exception) { Console.WriteLine("Ocorreu um erro ao finalizar o DTE mas o processamento prossegue"); }
            // Ultimate solution
            KillInstance();
        }

        public void KillInstance()
        {
            var activeVsApps = GetActiveVSApps();

            foreach (int vsApp in activeVsApps)
            {
                // When the vs instance was not in the list of active instances when the wrapper was constructed
                if (!ActiveVsApps.Exists((int processid) => processid == vsApp))
                {
                    // Try to kill the process (occurs when the devenv process hangs, or the operation has timed out, or something unexpected happened.
                    try
                    {
                        System.Diagnostics.Process.GetProcessById(vsApp).Kill();
                    }
                    catch (ArgumentException)
                    {
                        // Process is not running
                    }
                    catch (System.ComponentModel.Win32Exception)
                    {
                        // Access is denied
                    }
                }
            }
        }

        /// <summary>
        /// Exports the code metrics results to Excel.
        /// </summary>
        /// <returns>Returns the process id of the Excel process that contains the code metric results.</returns>
        public int ExportCodeMetrics()
        {
            int retries;

            object customin = null;
            object customout = null;

            Console.WriteLine("Exportando métricas");
            var startCalcMetrics = DateTime.Now;
            // Get the list of open Excel applications, so you know which excel application is created for the code metrics.
            var activeExcelApps = GetActiveExcelApps();

            // Start the export command
            Command exportMetricsCommand = null;
            retries = 0;
            while (true)
            {
                try
                {
                    exportMetricsCommand = Dte.Commands.Item("{79989DD6-4C13-4D10-9872-73538668D037}", 1287);
                    if (exportMetricsCommand.IsAvailable) break;

                    Console.Write(".");
                    System.Threading.Thread.Sleep(1000);
                }
                catch (COMException)
                {
                    if (retries < MaxRetries)
                    {
                        Console.Write(".");
                        System.Threading.Thread.Sleep(5000);
                        retries += 1;
                    }
                    else
                    {
                        ExportCodeMetricsResult = -1;
                        return -1;
                    }
                }
            }

            retries = 0;
            while (true)
            {
                try
                {
                    exportMetricsCommand.DTE.Commands.Raise("{79989DD6-4C13-4D10-9872-73538668D037}", 1287, ref customin, ref customout);
                    ((Excel.Application)Marshal.GetActiveObject("Excel.Application")).Visible=false;
                    
                    //Dte.Commands.Raise("{79989DD6-4C13-4D10-9872-73538668D037}", 1287, ref customin, ref customout);
                    break;
                }
                catch (COMException)
                {
                    System.Threading.Thread.Sleep(5000);
                    if (retries < MaxRetries)
                        retries += 1;
                    else
                    {
                        ExportCodeMetricsResult = -1;
                        return -1;
                    }
                }
            }

            var timeElapsedToCalc = DateTime.Now - startCalcMetrics;
            Console.WriteLine("");
            Console.WriteLine("Tempo gasto calculando métricas: " + timeElapsedToCalc);

            while (true)
            {
                foreach (int excelApp in GetActiveExcelApps())
                {
                    if (!activeExcelApps.Contains(excelApp))
                    {
                        return excelApp;
                    }
                }
            }
        }

        /// <summary>
        /// Look in the process list for all processes that are called 'excel'
        /// </summary>
        /// <returns>A list of excel process id's.</returns>
        private List<int> GetActiveVSApps()
        {
            List<int> activeVSApps = new List<int>();

            foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcesses())
            {
                if (proc.ProcessName.ToLower() == "devenv")
                {
                    activeVSApps.Add(proc.Id);
                }
            }

            return activeVSApps;
        }

        /// <summary>
        /// Look in the process list for all processes that are called 'excel'
        /// </summary>
        /// <returns>A list of excel process id's.</returns>
        private List<int> GetActiveExcelApps()
        {
            List<int> activeExcelApps = new List<int>();

            foreach (System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcesses())
            {
                if (proc.ProcessName.ToLower() == "excel")
                {
                    activeExcelApps.Add(proc.Id);
                }
            }

            return activeExcelApps;
        }

    }
}
