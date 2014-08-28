using System;
using System.Diagnostics;
using System.ServiceProcess;
using SMSSender;
using SMS_Sales_Service.Entities;
using System.Timers;
using Npgsql;
using SMSSender.DAO;

namespace SMS_Sales_Service
{
    public partial class SMS_Sales_Service : ServiceBase
    {
        Timer timer = new Timer();
        public SMS_Sales_Service()
        {
            InitializeComponent();
            AutoLog = false;
            if (!EventLog.SourceExists("SMSSalesSource"))
            {
                EventLog.CreateEventSource("SMSSalesSource", "SMS Sales Service");
            }
            myEventLog.Source = "SMSSalesSource";
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            try
            {
                //myEventLog.WriteEntry("Timer tick.", EventLogEntryType.Information);
                timer.Stop();

                ConnectionManager connectionManager = new ConnectionManager();
                NpgsqlConnection conn = connectionManager.getConnection();

                SettingsDAO settingsDAO = new SettingsDAO();
                Settings set = settingsDAO.readAll(conn);
                myEventLog.WriteEntry(set.ToString(),EventLogEntryType.Information);
                conn.Close();
                conn.Dispose();

                if (validateSettings(set))
                {

                    if (!DayEnabled(set))
                    {
                        timer.Start();
                        return;
                    }
                    myEventLog.WriteEntry("Inicia Proceso.", EventLogEntryType.Information);

                    resetWeek(set);
                    string sScheduled = set.todSchedule;
                    string[] scheduled = sScheduled.Split(';');
                    string current = DateTime.Now.ToString("HH:mm");
                    bool sent = false;
                    foreach (string sTime in scheduled)
                    {
                        if (current.CompareTo(sTime) >= 0 && !HourExecuted(sTime, set))
                        {
                            if (!sent && InExecutionRange(sTime))
                            {
                                sendSMS(set);
                                //localFiles(set);
                                sent = true;
                            }
                            addExecutionToList(sTime, set);
                        }
                    }
                }
                else
                {
                    myEventLog.WriteEntry("Configuracion incorrecta. Corrija la configuracion de SMS y posteriormente reinicie el servicio.", EventLogEntryType.Error);
                    return;
                }

                timer.Start();

            }
            catch (Exception ex)
            {
                myEventLog.WriteEntry("Ha ocurrido un error: " + ex.Message + " " + ex.StackTrace, EventLogEntryType.Error);
            }
        }
        protected override void OnStart(string[] args)
        {
            try
            {
                timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
                timer.Interval = 60000;
                timer.Enabled = true;
            }
            catch (Exception ex)
            {
                myEventLog.WriteEntry("Error: " + ex.Message, EventLogEntryType.Error);
            }
        }

        protected override void OnStop()
        {
            myEventLog.WriteEntry("Se detuvo el servicio.",EventLogEntryType.Information);
        }
        private bool DayEnabled(Settings set)
        {
            DateTime today = DateTime.Today;

            switch (today.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    return bool.Parse(set.monSchedule);
                case DayOfWeek.Tuesday:
                    return bool.Parse(set.tueSchedule);
                case DayOfWeek.Wednesday:
                    return bool.Parse(set.wedSchedule);
                case DayOfWeek.Thursday:
                    return bool.Parse(set.thuSchedule);
                case DayOfWeek.Friday:
                    return bool.Parse(set.friSchedule);
                case DayOfWeek.Saturday:
                    return bool.Parse(set.satSchedule);
                case DayOfWeek.Sunday:
                    return bool.Parse(set.sunSchedule);
            }
            return false;
        }
        private void resetWeek(Settings set)
        {
            DateTime today = DateTime.Today;
            string executionBackup = "";
            switch (today.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    executionBackup = set.monExecuted;
                    break;
                case DayOfWeek.Tuesday:
                    executionBackup = set.tueExecuted;
                    break;
                case DayOfWeek.Wednesday:
                    executionBackup = set.wedExecuted;
                    break;
                case DayOfWeek.Thursday:
                    executionBackup = set.thuExecuted;
                    break;
                case DayOfWeek.Friday:
                    executionBackup = set.friExecuted;
                    break;
                case DayOfWeek.Saturday:
                    executionBackup = set.satExecuted;
                    break;
                case DayOfWeek.Sunday:
                    executionBackup = set.sunExecuted;
                    break;
            }

            set.monExecuted = today.DayOfWeek == DayOfWeek.Monday ? executionBackup : "";
            set.tueExecuted = today.DayOfWeek == DayOfWeek.Tuesday ? executionBackup : "";
            set.wedExecuted = today.DayOfWeek == DayOfWeek.Wednesday ? executionBackup : "";
            set.thuExecuted = today.DayOfWeek == DayOfWeek.Thursday ? executionBackup : "";
            set.friExecuted = today.DayOfWeek == DayOfWeek.Friday ? executionBackup : "";
            set.satExecuted = today.DayOfWeek == DayOfWeek.Saturday ? executionBackup : "";
            set.sunExecuted = today.DayOfWeek == DayOfWeek.Sunday ? executionBackup : "";

            set.Save(myEventLog);

        }
        private bool HourExecuted(string current, Settings set)
        {
            DateTime today = DateTime.Today;
            string sExecuted = "";
            string[] executed;
            int found = 0;

            switch (today.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    sExecuted = set.monExecuted;
                    break;
                case DayOfWeek.Tuesday:
                    sExecuted = set.tueExecuted;
                    break;
                case DayOfWeek.Wednesday:
                    sExecuted = set.wedExecuted;
                    break;
                case DayOfWeek.Thursday:
                    sExecuted = set.thuExecuted;
                    break;
                case DayOfWeek.Friday:
                    sExecuted = set.friExecuted;
                    break;
                case DayOfWeek.Saturday:
                    sExecuted = set.satExecuted;
                    break;
                case DayOfWeek.Sunday:
                    sExecuted = set.sunExecuted;
                    break;
            }
            executed = sExecuted.Split(';');
            found = Array.IndexOf(executed, current);
            return found >= 0;
        }
        private bool InExecutionRange(string sTime)
        {
            string sTolerance;
            string current = DateTime.Now.ToString("HH:mm");
            DateTime tolerance = DateTime.Parse(sTime);
            tolerance = tolerance.AddMinutes(15);

            sTolerance = tolerance.ToString("HH:mm");
            return sTolerance.CompareTo(current) >= 0;
        }
        private void addExecutionToList(string sTime, Settings set)
        {
            DateTime today = DateTime.Today;
            string executionBackup = "";
            //// Open App.Config of executable
            //System.Configuration.Configuration config =
            // ConfigurationManager.OpenExeConfiguration
            //            (ConfigurationUserLevel.None);
            switch (today.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    executionBackup = set.monExecuted;
                    set.monExecuted = executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime;
                    break;
                case DayOfWeek.Tuesday:
                    executionBackup = set.tueExecuted;
                    set.tueExecuted = executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime;
                    break;
                case DayOfWeek.Wednesday:
                    executionBackup = set.wedExecuted;
                    set.wedExecuted = executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime;
                    break;
                case DayOfWeek.Thursday:
                    executionBackup = set.thuExecuted;
                    set.thuExecuted = executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime;
                    break;
                case DayOfWeek.Friday:
                    executionBackup = set.friExecuted;
                    set.friExecuted = executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime;
                    break;
                case DayOfWeek.Saturday:
                    executionBackup = set.satExecuted;
                    set.satExecuted = executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime;
                    break;
                case DayOfWeek.Sunday:
                    executionBackup = set.sunExecuted;
                    set.sunExecuted = executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime;
                    break;
            }

            // Save the changes in App.config file.
            set.Save(myEventLog);

        }
         private bool validateSettings(Settings set)
        {
            if (set.sms_user.Trim() == "") return false;
            return true;
        }
        private void localFiles(Settings set)
        {
            myEventLog.WriteEntry("Intentando envio de SMS.", EventLogEntryType.Information);
            try
            {
                SalesPickerPostgresql sp = new SalesPickerPostgresql();

                sp.UserName = set.sms_user;
                sp.Password = set.sms_password;
                sp.SendMethod = SMSSender.CommonConstants.SMS_LOCAL;
                //sp.SendMethod = SMSSender.CommonConstants.SMS_MAS_MENSAJES;

                sp.RetrieveData();
                sp.SendSMS(myEventLog);
                sp.SendBossSMS(myEventLog);
            }
            catch (Exception e)
            {
                myEventLog.WriteEntry("Mensaje de Exception: " + e.Message, EventLogEntryType.Error);
                myEventLog.WriteEntry("Pila de Exception: " + e.StackTrace, EventLogEntryType.Error);
            }
            myEventLog.WriteEntry("Archivos locales de simulación de SMS generados exitosamente.", EventLogEntryType.Information);
        }
        private void sendSMS(Settings set)
        {
            myEventLog.WriteEntry("Intentando envio de SMS.", EventLogEntryType.Information);
            
            try
            {
                //Postgresql mode
                SalesPickerPostgresql sp = new SalesPickerPostgresql();
                sp.UserName = set.sms_user;
                sp.Password = set.sms_password;
                sp.SendMethod = SMSSender.CommonConstants.SMS_MAS_MENSAJES;
                sp.RetrieveData();
                sp.SendSMS(myEventLog);
                sp.SendBossSMS(myEventLog);
            }
            catch (Exception e)
            {
                myEventLog.WriteEntry("Mensaje de Excepcion: " + e.Message, EventLogEntryType.Error);
                myEventLog.WriteEntry("Pila de Excepcion: " + e.StackTrace, EventLogEntryType.Error);
            }
            myEventLog.WriteEntry("Proceso de envio de SMS finalizado exitosamente.", EventLogEntryType.SuccessAudit);
        }
        
    }
}
