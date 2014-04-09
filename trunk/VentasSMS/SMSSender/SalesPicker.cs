using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using CommonAdminPaq;
using CommonAdminPaq.dto;
using System.Runtime.InteropServices;
using System.Globalization;

namespace SMSSender
{
    class SalesPicker
    {
        private string workbookPath;
        private IList<Empresa> empresas;
        private string sendMethod;
        private ISendable sendable;
        private AdminPaqImpl api;

        public string WorkbookPath { get { return workbookPath; } set { workbookPath = value; } }
        public string SendMethod { get { return sendMethod; }
            set
            {
                switch (value) 
                {
                    case CommonConstants.SMS_MAS_MENSAJES:
                        sendable = new MasMensajes();
                        sendable.DoLogin("user", "password");
                        break;
                    case CommonConstants.SMS_LOCAL:
                        sendable = new LocalSMS();
                        break;
                    default:
                        sendable = new MasMensajes();
                        break;
                }
                sendMethod = value;
            }
        }

        # region EXCEL_CONFIG
        private Excel.Range RangoDeColumna(Excel.Worksheet workingSheet, int columnIndex, int startRow)
        {
            Object oCurrent;
            string startName, endName;
            int unicode = 64 + columnIndex;
            char character = (char)unicode;
            int currentRow = startRow;


            startName = character.ToString() + startRow;
            oCurrent = workingSheet.Cells[currentRow, columnIndex].Value;
            
            while(oCurrent != null)
            {
                currentRow++;
                oCurrent = workingSheet.Cells[currentRow, columnIndex].Value;
            }
            currentRow--;
            endName = character.ToString() + currentRow;
            return workingSheet.Range[startName + ":" + endName];
        }

        private bool codigoEnRango(Excel.Range rango, string valor)
        {
            foreach (Excel.Range currentRow in rango.Rows)
            {
                if (currentRow.Value == valor)
                    return true;
            }
            return false;
        }

        private void RetrieveSaleCodes(Empresa empresa, Excel.Range salesRange)
        {
            foreach (Excel.Range rSales in salesRange)
            {
                int currentRow = rSales.Row - 6;
                Excel.Range currentRange = salesRange.Cells[currentRow, 1];
                object oTemp = currentRange.Value;
                if (oTemp != null) {
                    string sTemp = oTemp.ToString();
                    empresa.CodigosVenta.Add(long.Parse(sTemp));
                }
            }
        }

        private void RetrieveReturnCodes(Empresa empresa, Excel.Range returnRange)
        {
            foreach (Excel.Range rSales in returnRange)
            {
                int currentRow = rSales.Row - 6;
                Excel.Range currentRange = returnRange.Cells[currentRow, 1];
                object oTemp = currentRange.Value;
                if (oTemp != null)
                {
                    string sTemp = oTemp.ToString();
                    empresa.CodigosDevolucion.Add(long.Parse(sTemp));
                }
            }
        }

        private void RetrieveBosses(Empresa empresa, Excel.Range bossRange)
        {
            foreach (Excel.Range rowRange in bossRange)
            {
                int currentRow = rowRange.Row - 6;
                Excel.Range currentRange = bossRange.Cells[currentRow, 1];
                object oTemp = currentRange.Value;
                if (oTemp != null)
                {
                    string sTemp = oTemp.ToString();
                    empresa.TelefonosDirectores.Add(sTemp);
                }
            }
        }

        private void RetrieveConfiguration(Excel.Worksheet workingSheet)
        {
            Excel.Range rangoCodigosDevolucion, rangoCodigosVenta, rangoCodigoAgentes, rangoTelefonos, rangoMeta, rangoSysAgente, rangoDirectores;
            rangoCodigosDevolucion = RangoDeColumna(workingSheet, 2, 7);
            rangoCodigosVenta = RangoDeColumna(workingSheet, 1, 7);
            rangoDirectores = RangoDeColumna(workingSheet, 11, 7);
            
            rangoCodigoAgentes = RangoDeColumna(workingSheet, 4, 7);
            int lastAgenteRow = 7 + rangoCodigoAgentes.Rows.Count - 1;

            rangoMeta = workingSheet.Range["F7:F" + lastAgenteRow];
            rangoTelefonos = workingSheet.Range["G7:G" + lastAgenteRow];
            rangoSysAgente = workingSheet.Range["H7:H" + lastAgenteRow];

            Empresa empresaHoja = new Empresa();
            object oTemp = null;
            
            oTemp = workingSheet.Range["B1"].Value;

            // If no Id found for company, it has not been validated.
            if (oTemp == null) return;
            empresaHoja.Id = long.Parse(oTemp.ToString());

            oTemp = workingSheet.Range["B2"].Value;
            empresaHoja.Nombre = oTemp == null?"":oTemp.ToString();

            oTemp = workingSheet.Range["B3"].Value;
            empresaHoja.Ruta = oTemp == null?"":oTemp.ToString();

            RetrieveSaleCodes(empresaHoja, rangoCodigosVenta);
            RetrieveReturnCodes(empresaHoja, rangoCodigosDevolucion);
            RetrieveBosses(empresaHoja, rangoDirectores);

            IList<Agente> sheetAgents = new List<Agente>();
            // Retrieve agentes from worksheet
            foreach(Excel.Range rTel in rangoTelefonos)
            {
                int currentRow = rTel.Row - 6;
                Excel.Range rAgente = rangoSysAgente.Cells[currentRow, 1];
                object oSysId = rAgente.Value;
                if (oSysId == null) continue;

                Agente rowAgent = new Agente();
                rowAgent.Id = long.Parse(oSysId.ToString());

                Excel.Range rCAgente = rangoCodigoAgentes.Cells[currentRow, 1];
                oTemp = rCAgente.Value;
                rowAgent.Codigo = oTemp == null ? "" : oTemp.ToString();

                Excel.Range rPhone = rangoTelefonos.Cells[currentRow, 1];
                object oPhone = rPhone.Value;
                if (oPhone == null) continue;

                CellPhone RowPhone = new CellPhone();
                RowPhone.PhoneNumber = oPhone.ToString();

                Excel.Range rMeta = rangoMeta.Cells[currentRow, 1];
                object oMeta = rMeta.Value;
                if (oMeta != null) 
                {
                    RowPhone.MetaSemanal = double.Parse(oMeta.ToString());
                }
                

                rowAgent.Phone = RowPhone;
                sheetAgents.Add(rowAgent);
            }

            empresaHoja.Agentes = sheetAgents;
            Console.WriteLine("found " + empresaHoja.Agentes.Count + " Sales agents in the curent worksheet.");
            empresas.Add(empresaHoja);
        }

        public void RetrieveConfiguration() 
        {
            if (!validateMe()) return;

            Excel.Application xlApp;
            Excel.Workbook book;

            Console.WriteLine("Working with excel file: " + workbookPath);
            xlApp = new Excel.Application();
            xlApp.Visible = false;
            book = xlApp.Workbooks.Open(workbookPath);

            empresas = new List<Empresa>();

            // Implementation for info gathering here.
            foreach (Excel.Worksheet workingSheet in book.Worksheets)
            {
                string companyName = workingSheet.Range["B2"].Value;
                Console.WriteLine("Working in worksheet [" + workingSheet.Name + "] with company name [" + companyName + "]");
                RetrieveConfiguration(workingSheet);
            }

            Console.WriteLine("Releasing excel resources");
            book.Close();
            xlApp.Quit();

            Marshal.FinalReleaseComObject(book);
            Marshal.FinalReleaseComObject(xlApp);
        }

        # endregion



        public void RetrieveSales() {
            if (!validateMe() || empresas == null) return;
            api = new AdminPaqImpl();
            api.InitializeSDK();

            foreach(Empresa empresa in empresas)
            {
                Console.WriteLine("Retrieving sales for company: [" + empresa.Nombre + "]");
                api.RetrieveSales(empresa);
            }
        }

        public void SendSMS() {
            if (!validateMe() || empresas == null) return;
            foreach (Empresa empresa in empresas)
            {
                string sms;
                foreach (Agente agente in empresa.Agentes)
                {
                    CultureInfo provider = CultureInfo.CreateSpecificCulture("en-US");

                    string cumplimientoSemanal = agente.Phone.CumplimientoSemana.ToString("p");
                    string cumplimientoTentencia = agente.Phone.CumplimientoTendencia.ToString("p");
                    string faltanteSemanal = agente.Phone.FaltanteSemana.ToString("C", provider);

                    empresa.ResultadoSemanal += string.Format("{0}:{1};", agente.Codigo, cumplimientoSemanal);
                    empresa.Tendencia += string.Format("{0}:{1};", agente.Codigo, cumplimientoTentencia);
                    //string faltanteTendencia = agente.Phone.FaltanteTendencia.ToString("C", provider);

                    string semanal = string.Format("Tu meta de venta de esta semana ha sido cumplida en un {0}, faltan {1} por vender",
                        cumplimientoSemanal, faltanteSemanal);
                    if (agente.Phone.CumplimientoSemana >= 1)
                        semanal = "Felicidades! Haz alcanzado el {0} de ventas en tu meta semanal";

                    string tendencia = string.Format("Tendencia de cumplimiento a 4 semanas {0}", cumplimientoTentencia);


                    sms = string.Format("{0}. {1}", semanal, tendencia);
                    byeSMS(agente.Phone.PhoneNumber, sms);
                }

            }
        }

        public void SendBossSMS()
        {

            foreach (Empresa empresa in empresas)
            {
                string tituloSemanal, tituloTendencia, sms;

                tituloSemanal = "Cumplimiento de metas en venta semanal";
                sms = string.Format("{0}: {1}", tituloSemanal, empresa.ResultadoSemanal);
                SendBossSMS(sms, empresa);
                
                tituloTendencia = "Tendencia de cumplimiento en metas de ventas a 4 semanas";
                sms = string.Format("{0}: {1}", tituloTendencia, empresa.Tendencia);
                SendBossSMS(sms, empresa);
            }
        }

        private void SendBossSMS(string sms, Empresa empresa)
        { 
            foreach(string boss in empresa.TelefonosDirectores)
            {
                byeSMS(boss, sms);
            }
        }

        private void byeSMS(string to, string msg)
        {
            Sms sms = new Sms();
            sms.To = to;
            sms.Message = msg;
            sendable.SendSMS(sms);
        }

        private bool validateMe() {
            if ("".Equals(workbookPath))
            {
                ErrLogger.Log("SalesPicker - USAGE ERROR: MUST SET WORKBOOK PATH");
                Console.WriteLine("SalesPicker - USAGE ERROR: MUST SET WORKBOOK PATH");
                return false;
            }

            if (sendable == null)
            {
                ErrLogger.Log("SalesPicker - USAGE ERROR: DID NOT SET A SENDABLE");
                Console.WriteLine("SalesPicker - USAGE ERROR: DID NOT SET A SENDABLE");
                return false;
            }
            return true;
        }
    }
}
