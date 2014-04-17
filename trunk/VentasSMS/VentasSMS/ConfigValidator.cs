using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommonAdminPaq;
using Excel = Microsoft.Office.Interop.Excel;
using CommonAdminPaq.dto;
using System.Runtime.InteropServices;

namespace VentasSMS
{
    public class ConfigValidator
    {
        private string ruta;
        private IList<string> errors = new List<string>();
        private AdminPaqImpl api;
        
        public string Ruta { get { return ruta; } set { ruta = value; } }
        public IList<string> ValidationErrors { get { return errors; } set { errors = value; } }
        public AdminPaqImpl API { get { return api; } set { api = value; } }

        public bool ValidateConfiguration() 
        {
            return validateFile();
        }

        private bool validateFile()
        {
            bool validated = false;
            int sheetIdx = 1;
            // Validate each worksheet

            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = false;

            Excel.Workbook book = xlApp.Workbooks.Open(ruta);

            while (sheetIdx <= book.Worksheets.Count)
            {
                validated = validateWorkSheet(book.Worksheets[sheetIdx]);
                if (!validated) 
                {
                    validated = false;
                    break;
                } 
                sheetIdx++;
            }

            if (validated) book.Save();

            book.Close(SaveChanges: false);
            xlApp.Quit();

            Marshal.FinalReleaseComObject(book);
            Marshal.FinalReleaseComObject(xlApp);

            return validated;
        }

        private bool validateWorkSheet(Excel.Worksheet evalSheet)
        {
            // Validate
            if (!ValidarEmpresa(evalSheet))
            {
                evalSheet.Range["D1"].Value = "";
                return false;
            }

            if (!ValidarCodigosVenta(evalSheet))
            {
                evalSheet.Range["D1"].Value = "";
                return false;
            }

            if (!ValidarCodigosDevolucion(evalSheet))
            {
                evalSheet.Range["D1"].Value = "";
                return false;
            }

            if (!ValidarAgentesDeVenta(evalSheet))
            {
                evalSheet.Range["D1"].Value = "";
                return false;
            }

            evalSheet.Range["D1"].Value = "sysvalidated";
            return true;
        }

        private bool ValidarEmpresa(Excel.Worksheet evalSheet)
        {
            string sEmpresa = evalSheet.Range["B2"].Value;

            bool empresaEnLista = false;
            foreach (Empresa empresa in api.Empresas)
            {
                if (sEmpresa.ToLower().Equals(empresa.Nombre.ToLower()))
                {
                    empresaEnLista = true;
                    evalSheet.Range["B1"].Value = empresa.Id;
                    evalSheet.Range["B3"].Value = empresa.Ruta;
                    break;
                }
            }
            if (!empresaEnLista)
            {
                ErrLogger.Log("Company not found in database: " + sEmpresa);
            }

            return empresaEnLista;
        }


        private bool ValidarCodigosVenta(Excel.Worksheet evalSheet)
        {
            Excel.Range salesRange = RangoDeColumna(evalSheet, 1, 7);
            foreach (Excel.Range rSales in salesRange)
            {
                int currentRow = rSales.Row - 6;
                Excel.Range currentRange = salesRange.Cells[currentRow, 1];
                object oTemp = currentRange.Value;
                if (oTemp != null)
                {
                    string sTemp = oTemp.ToString();
                    if (!api.CodigoDocoValido(sTemp, evalSheet.Range["B3"].Value))
                    {
                        ErrLogger.Log("Sales Document Code found to be invalid in database: [" + sTemp + "]");
                        return false;
                    }
                }
            }
            return true;
        }

        private bool ValidarCodigosDevolucion(Excel.Worksheet evalSheet)
        {
            Excel.Range salesRange = RangoDeColumna(evalSheet, 2, 7);
            foreach (Excel.Range rSales in salesRange)
            {
                int currentRow = rSales.Row - 6;
                Excel.Range currentRange = salesRange.Cells[currentRow, 1];
                object oTemp = currentRange.Value;
                if (oTemp != null)
                {
                    string sTemp = oTemp.ToString();
                    if (!api.CodigoDocoValido(sTemp, evalSheet.Range["B3"].Value))
                    {
                        ErrLogger.Log("Refunds Document Code found to be invalid in database: [" + sTemp + "]");
                        return false;
                    }
                }
            }
            return true;
        }

        private bool ValidarAgentesDeVenta(Excel.Worksheet evalSheet)
        {
            Excel.Range salesRange = RangoDeColumna(evalSheet, 4, 7);
            foreach (Excel.Range rSales in salesRange)
            {
                int currentRow = rSales.Row - 6;
                Excel.Range currentRange = salesRange.Cells[currentRow, 1];
                object oTemp = currentRange.Value;
                if (oTemp != null)
                {
                    string sTemp = oTemp.ToString();
                    int sysId = api.IdAgente(sTemp, evalSheet.Range["B3"].Value);
                    if (sysId == 0)
                    {
                        ErrLogger.Log("Sales Agent Code found to be invalid in database: [" + sTemp + "]");
                        return false;
                    }
                    else
                    {
                        evalSheet.Range["H" + rSales.Row].Value = sysId;
                    }
                }
            }
            return true;
        }

        private Excel.Range RangoDeColumna(Excel.Worksheet workingSheet, int columnIndex, int startRow)
        {
            Object oCurrent;
            string startName, endName;
            int unicode = 64 + columnIndex;
            char character = (char)unicode;
            int currentRow = startRow;


            startName = character.ToString() + startRow;
            oCurrent = workingSheet.Cells[currentRow, columnIndex].Value;

            while (oCurrent != null)
            {
                currentRow++;
                oCurrent = workingSheet.Cells[currentRow, columnIndex].Value;
            }
            currentRow--;
            endName = character.ToString() + currentRow;
            return workingSheet.Range[startName + ":" + endName];
        }

        
    }
}
