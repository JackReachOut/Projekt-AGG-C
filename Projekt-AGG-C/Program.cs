using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, string sheetName)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            try
            {
                ws = wb.Worksheets[sheetName];
                ExcelFound = true;
            }
            catch
            {
                ExcelFound = false;
            }
        }


        public string ReadCell(int i, int j)
        {
            if (ws.Cells[i, j].Value2 != null)
                return ws.Cells[i, j].Value2.ToString();
            else
                return "";
        }

        public void quitExcel()
        {
            // Save and close workbook
            wb.Save();
            wb.Close();

            // Quit Excel application
            excel.Quit();

            // Release COM objects to prevent memory leaks
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

            // Kill Excel process
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }

        public bool ExcelFound { get; set; }
    }



    class Program
    {
        static void Main(string[] args)
        {
            Excel excel = new Excel(@"Verzeichnis", "Tabelle2");

            if (excel.ExcelFound)
            {
                string bestandB2 = excel.ReadCell(16, 16);
                string bestandBS = excel.ReadCell(17, 16);
                string umlaufB = excel.ReadCell(17, 17);
                string bestandT = excel.ReadCell(18, 18);
                string umlaufT = excel.ReadCell(19, 19);
                string bestandF = excel.ReadCell(15, 15);

                if (int.TryParse(bestandB2, out int bestandB2Int) && int.TryParse(bestandBS, out int bestandBSInt))
                {
                    int bestandBeGe = bestandB2Int + bestandBSInt;
                    Console.WriteLine("Bestand B2: " + bestandB2Int);
                    Console.WriteLine("Bestand BS: " + bestandBSInt);
                    Console.WriteLine("Bestand BeGe: " + bestandBeGe);

                    // write to Meeting_Board.xlsx, Sheet3
                    _Excel.Application excelApp = new _Excel.Application();
                    Workbook wb = excelApp.Workbooks.Open(@"Verzeichnis.xlsx");
                    Worksheet ws = wb.Worksheets[3];

                    ws.Cells[1, 1] = bestandBeGe;
                    ws.Cells[2, 2] = umlaufB;
                    ws.Cells[3, 3] = bestandT;

                    wb.Save();
                    wb.Close();

                    excel.quitExcel();
                }
                else
                {
                    Console.WriteLine("Could not parse one or more values as integers.");
                }

                Console.WriteLine("Umlauf B: " + umlaufB);
                Console.WriteLine("Bestand T: " + bestandT);
                Console.WriteLine("Umlauf T: " + umlaufT);
                Console.WriteLine("Bestand F: " + bestandF);

                //excel.quitExcel();
            }
            else
            {
                Console.WriteLine("Excel-Datei oder Worksheet konnte nicht gefunden werden!");
            }

            Console.ReadLine();
        }
    }
}

