using System;
using System.Collections.Generic;
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

        private List<Enterprise> listEnterprises = new List<Enterprise>();
        private List<Seller> listSellers = new List<Seller>();
        private List<Director> listDirectors =new List<Director>();


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

            sellerDAO.setEnterprisesToSellers(enterpriseDAO);
            directorDAO.setEnterprisesToDirectors(enterpriseDAO);

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
                foreach (Seller agente in empresa.Agentes)
                {
                    if (agente.WeeklyGoal > 0)
                    {
                        CultureInfo provider = CultureInfo.CreateSpecificCulture("en-US");

                        float cumplimientoSemanal =  agente.CumplimientoSemana / agente.WeeklyGoal;
                        float faltanteSemanal = agente.WeeklyGoal - agente.CumplimientoSemana;
                        float cumplimientoMensual = agente.CumplimientoMensual / (agente.WeeklyGoal * 4);
                        float faltanteMensual = agente.WeeklyGoal * 4 - agente.CumplimientoMensual;

                        empresa.ResultadoSemanal += string.Format("{0}: {1}; ", agente.Code, cumplimientoSemanal.ToString("p"));
                        empresa.ResultadoMensual += string.Format("{0}: {1}; ", agente.Code, cumplimientoMensual.ToString("p"));
                        //string faltanteTendencia = agente.Phone.FaltanteTendencia.ToString("C", provider);

                        string semanal = string.Format("Tu meta de venta semanal es cumplida en {0}, falta {1} por vender.",
                            cumplimientoSemanal.ToString("p"), faltanteSemanal.ToString("C", provider));

                        if (agente.CumplimientoSemana >= 1)
                            semanal = string.Format("Felicidades! Haz alcanzado el {0} de ventas en tu meta semanal.", cumplimientoSemanal.ToString("p"));

                        string mensual = string.Format("Tu meta de venta mensual es cumplida en {0}, faltan {1} por vender.",
                            cumplimientoMensual.ToString("p"), faltanteMensual.ToString("C", provider));

                        if (agente.CumplimientoMensual >= 1)
                            mensual = string.Format("Felicidades! Haz alcanzado el {0} de ventas en tu meta mensual.", cumplimientoMensual.ToString("p"));

                        sms = string.Format("{0}" + Environment.NewLine + "{1}", semanal, mensual);
                        byeSMS(agente.CellPhone, sms, userName, password);
                    }
                }
            }
        }

        public void SendBossSMS()
        {
            foreach (Enterprise empresa in listEnterprises)
            {   
                string tituloSemanal, tituloMensual, sms;
                if (empresa.ResultadoSemanal.Trim() != "")
                {
                    tituloSemanal = "Cumplimiento de metas en venta semanal";
                    sms = string.Format("{0}: {1}", tituloSemanal, empresa.ResultadoSemanal);
                    SendBossSMS(sms, empresa);
                }
                if (empresa.ResultadoMensual.Trim() != "")
                {
                    tituloMensual = "Cumplimiento de metas en venta mensual";
                    sms = string.Format("{0}: {1}", tituloMensual, empresa.ResultadoMensual);
                    SendBossSMS(sms, empresa);
                }
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
