using System;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;

namespace CodeMetricCalculator
{
    public class ExcelHandler : IDisposable
    {
        private int ProcessId { get; set; }
        private Excel.Application Application { get; set; }
        private Excel.Workbook Workbook { get; set; }
        private Excel.Worksheet Worksheet { get; set; }

        public ExcelHandler(int processId)
        {
            this.ProcessId = processId;
            // Kill all other Excel instances
            KillOtherExcelInstances(processId);

            // Get the application that contains the export results
            int triesExcel = 0;
            ;
            while (true)
            {
                try
                {
                    Application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    Application.Visible = false;
                    break;
                }
                catch (COMException ex)
                {
                    if (triesExcel <= 10)
                    {
                        triesExcel++;
                        Thread.Sleep(1000);
                    }
                    else
                    Console.WriteLine("   Failed to activate Excel ...");
                }
            }
            Workbook = Application.Workbooks[1];
            Worksheet = (Excel.Worksheet)Workbook.Worksheets[1];
        }
        public void Dispose()
        {
            if (Process.GetProcessById(ProcessId) != null)
                Process.GetProcessById(ProcessId).Kill();
        }

        private void KillOtherExcelInstances(int processId)
        {
            foreach (Process p in Process.GetProcessesByName("excel"))
            {
                if (p.Id == processId)
                    continue;

                p.Kill();
            }
        }

        internal void SaveResult(string outputFilePath, string fileName)
        {
            try
            {
                if (!Directory.Exists(outputFilePath))
                {
                    Directory.CreateDirectory(outputFilePath);
                }

                string fullPath = outputFilePath + fileName + ".xlsx";
                Workbook.SaveCopyAs(fullPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ocorreu um erro ao tentar salvar a planilha.");
            }
        }
    }
}
