using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using CommonAdminPaq;
using CommonAdminPaq.dto;
using System.Runtime.InteropServices;
using System.Globalization;
using Npgsql;
using SMSSender.Entities;
using SMSSender.DAO;

namespace SMSSender
{
    public class SalesPickerPostgresql
    {
        private string sendMethod, userName, password;
        private ISendable sendable;

        private List<Enterprise> listEnterprises;
        private List<Seller> listSellers;
        private List<Director> listDirectors;


        public string UserName { get { return userName; } set { userName = value; } }
        public string Password { get { return password; } set { password = value; } }
        public string SendMethod { get { return sendMethod; }
            set
            {
                switch (value) 
                {
                    case CommonConstants.SMS_MAS_MENSAJES:
                        sendable = new MasMensajes();
                        sendable.DoLogin(userName, password);
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

        public void RetrieveData()
        {
            ConnectionManager connectionManager = new ConnectionManager();

            SellerDAO sellerDAO = new SellerDAO();
            EnterpriseDAO enterpriseDAO = new EnterpriseDAO();
            DirectorDAO directorDAO = new DirectorDAO();

            NpgsqlConnection conn = connectionManager.getConnection();

            enterpriseDAO.readAll(conn);
            sellerDAO.readAll(conn);
            directorDAO.readAll(conn);

            sellerDAO.setEnterprisesToSellers();
            directorDAO.setEnterprisesToDirectors();

            listEnterprises = sellerDAO.readEmpresasInUse(listEnterprises);
            listEnterprises = directorDAO.readEmpresasInUse(listEnterprises);

            listDirectors = directorDAO.listDirectors;
            listSellers = sellerDAO.listSellers;

            conn.Close();
            conn.Dispose();
        }

        public void SendSMS() {

            if (!validateMe()) return;


            foreach (Enterprise empresa in listEnterprises)
            {
                empresa.ResultadoSemanal = "";
                empresa.ResultadoMensual = "";
                string sms;
                foreach (Seller agente in listSellers)
                {
                    CultureInfo provider = CultureInfo.CreateSpecificCulture("en-US");

                    float cumplimientoSemanal = 100 * agente.CumplimientoSemana / agente.WeeklyGoal;
                    float faltanteSemanal = agente.WeeklyGoal - agente.CumplimientoSemana;
                    float cumplimientoMensual = 100 * agente.CumplimientoMensual / (agente.WeeklyGoal * 4);
                    float faltanteMensual = agente.WeeklyGoal * 4 - agente.CumplimientoMensual;

                    empresa.ResultadoSemanal += string.Format("{0}:{1};", agente.Code, cumplimientoSemanal);
                    empresa.ResultadoMensual += string.Format("{0}:{1};", agente.Code, cumplimientoMensual);
                    //string faltanteTendencia = agente.Phone.FaltanteTendencia.ToString("C", provider);

                    string semanal = string.Format("Tu meta de venta de esta semana ha sido cumplida en un {0}, faltan {1} por vender",
                        cumplimientoSemanal.ToString("p"), faltanteSemanal.ToString("C", provider));

                    if (agente.CumplimientoSemana >= 1)
                        semanal = string.Format("Felicidades! Haz alcanzado el {0} de ventas en tu meta semanal", cumplimientoSemanal.ToString("p"));

                    string mensual = string.Format("Tu meta de venta de este mes ha sido cumplida en un {0}, faltan {1} por vender",
                        cumplimientoMensual.ToString("p"), faltanteMensual.ToString("C", provider));

                    if (agente.CumplimientoMensual >= 1)
                        mensual = string.Format("Felicidades! Haz alcanzado el {0} de ventas en tu meta semanal", cumplimientoMensual.ToString("p"));

                    sms = string.Format("{0}. {1}", semanal, mensual);
                    byeSMS(agente.CellPhone, sms, userName, password);
                }
            }
        }

        public void SendBossSMS()
        {
            foreach (Enterprise empresa in listEnterprises)
            {
                string tituloSemanal, tituloMensual, sms;

                tituloSemanal = "Cumplimiento de metas en venta semanal";
                sms = string.Format("{0}: {1}", tituloSemanal, empresa.ResultadoSemanal);
                SendBossSMS(sms, empresa);
                
                tituloMensual = "Cumplimiento de metas en venta mensual";
                sms = string.Format("{0}: {1}", tituloMensual, empresa.ResultadoMensual);
                SendBossSMS(sms, empresa);
            }
        }

        private void SendBossSMS(string sms, Enterprise empresa)
        { 
            foreach(Director boss in empresa.Directors)
            {
                byeSMS(boss.CellPhone, sms, userName, password);
            }
        }

        private void byeSMS(string to, string msg, string user, string pwd)
        {
            Sms sms = new Sms();
            sms.To = to;
            sms.Message = msg;
            sendable.DoLogin(user, pwd);
            sendable.SendSMS(sms);
        }

        private bool validateMe() {
            //if ("".Equals(workbookPath))
            //{
            //    ErrLogger.Log("SalesPicker - USAGE ERROR: MUST SET WORKBOOK PATH");
            //    Console.WriteLine("SalesPicker - USAGE ERROR: MUST SET WORKBOOK PATH");
            //    return false;
            //}

            //if (sendable == null)
            //{
            //    ErrLogger.Log("SalesPicker - USAGE ERROR: DID NOT SET A SENDABLE");
            //    Console.WriteLine("SalesPicker - USAGE ERROR: DID NOT SET A SENDABLE");
            //    return false;
            //}
            return true;
        }
    }
}
