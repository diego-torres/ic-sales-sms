using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using CommonAdminPaq;
using SMSSender;
using VentasSMS.Properties;

namespace VentasSMS
{
    public partial class frmMain : Form
    {
        private AdminPaqImpl api;
        public AdminPaqImpl API { get { return api; } }
        private frmConfig fConfig = null;
        private ScheduleConfig sConfig = null;
        private Auditoria Ventana_Auditoria = null;

        public frmMain()
        {
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            api = new AdminPaqImpl();
            api.InitializeSDK();
            timer1.Start();
        }

        private void openConfigForm(string fileName) 
        {
            if (fConfig == null || fConfig.IsDisposed)
            {
                fConfig = new frmConfig();
                fConfig.API = api;
                fConfig.textBoxFileName.Text = fileName;
            }

            fConfig.Show();
        }

        

        private void openScheduleConfig()
        {
            if (sConfig == null || sConfig.IsDisposed)
            {
                sConfig = new ScheduleConfig();
            }

            sConfig.LoadConfiguredSchedule();
            sConfig.Show();
        }

        private void openAuditoria()
        { 
            if (Ventana_Auditoria==null|| Ventana_Auditoria.IsDisposed)
            {
                Ventana_Auditoria = new Auditoria();
            }
            Ventana_Auditoria.Show();
        }   

        private string configuredFilePath()
        {
            Settings set = Settings.Default;

            string configFilePath = set.smsWorkbook;
            string myDocs = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            return configFilePath.Replace("~", myDocs);
        }

        private void telefonosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // EVALUATE IF CONFIGURED FILE EXISTS
            // For read access you do not need to call OpenExeConfiguraton
            string configFilePath = configuredFilePath();

            if ("".Equals(configFilePath.Trim()) || !File.Exists(configFilePath))
            {
                openConfigForm(configFilePath);
            }
            else 
            {
                System.Diagnostics.Process.Start(configFilePath);
            }

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.notifyIconSMS.Visible = true;
            this.Hide();
            timer1.Stop();
            timerSMS.Start();
        }

        private void scheduleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openScheduleConfig();
        }

        private void resetWeek()
        {
            Settings set = Settings.Default;
            DateTime today = DateTime.Today;
            string executionBackup="";
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

            set.Save();

            // Force a reload of a changed section.
            set.Reload();
        }

        private bool HourExecuted (string current)
        {
            Settings set = Settings.Default;
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
            return found>=0;
        }

        private void addExecutionToList(string sTime)
        {
            Settings set = Settings.Default;
            DateTime today = DateTime.Today;
            string executionBackup = "";
            // Open App.Config of executable
            System.Configuration.Configuration config =
             ConfigurationManager.OpenExeConfiguration
                        (ConfigurationUserLevel.None);
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
            set.Save();

            // Force a reload of a changed section.
            set.Reload();

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

        private void localFiles()
        {
            string configFilePath = configuredFilePath();
            ConfigValidator cv = new ConfigValidator();
            cv.API = api;
            cv.Ruta = configFilePath;

            if ("".Equals(configFilePath.Trim()) || !File.Exists(configFilePath) || !cv.ValidateConfiguration())
            {
                openConfigForm(configFilePath);
            }
            else
            {
                try
                {
                    Settings set = Settings.Default;
                    SalesPicker sp = new SalesPicker();
                    sp.WorkbookPath = configFilePath;
                    sp.UserName = set.user;
                    sp.Password = set.password;
                    sp.SendMethod = SMSSender.CommonConstants.SMS_LOCAL;
                    //sp.SendMethod = SMSSender.CommonConstants.SMS_MAS_MENSAJES;

                    sp.RetrieveConfiguration();
                    sp.RetrieveSales();
                    sp.SendSMS();
                    sp.SendBossSMS();
                }
                catch (Exception e)
                {
                    ErrLogger.Log(e.Message);
                    ErrLogger.Log(e.StackTrace);
                }

                MessageBox.Show("Archivos locales de simulación de SMS generados exitosamente.", "Simulación de SMS Local", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

        private void sendSMS()
        {
            string configFilePath = configuredFilePath();
            ConfigValidator cv = new ConfigValidator();
            cv.API = api;
            cv.Ruta = configFilePath;
            
            if ("".Equals(configFilePath.Trim()) || !File.Exists(configFilePath)||!cv.ValidateConfiguration())
            {
                openConfigForm(configFilePath);
            }
            else
            {
                try {
                    Settings set = Settings.Default;
                    SalesPicker sp = new SalesPicker();
                    sp.WorkbookPath = configFilePath;
                    sp.UserName = set.user;
                    sp.Password = set.password;
                    //sp.SendMethod = SMSSender.CommonConstants.SMS_LOCAL;
                    sp.SendMethod = SMSSender.CommonConstants.SMS_MAS_MENSAJES;

                    sp.RetrieveConfiguration();
                    sp.RetrieveSales();
                    sp.SendSMS();
                    sp.SendBossSMS();
                }catch(Exception e){
                    ErrLogger.Log(e.Message);
                    ErrLogger.Log(e.StackTrace);
                }
            }
        }

        private bool DayEnabled()
        {
            Settings set = Settings.Default;
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

        private void timerSMS_Tick(object sender, EventArgs e)
        {
            Settings set = Settings.Default;
            timerSMS.Stop();
            if (!DayEnabled()) 
            {
                timerSMS.Start();
                return;
            }
            

            resetWeek();
            string sScheduled = set.todSchedule;
            string[] scheduled = sScheduled.Split(';');
            string current = DateTime.Now.ToString("HH:mm");
            bool sent = false;
            foreach(string sTime in scheduled)
            {
                if (current.CompareTo(sTime) >= 0 && !HourExecuted(sTime))
                {
                    if (!sent && InExecutionRange(sTime))
                    {
                        sendSMS();
                        sent = true;
                    }
                    addExecutionToList(sTime);
                }
            }
            timerSMS.Start();
        }

        private void enviarSMSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult drSend = MessageBox.Show("¿Desea usted enviar los resúmenes de venta por SMS ahora?", "Confirmación de envío", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if(drSend == DialogResult.Yes)
                sendSMS();
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult drClose = MessageBox.Show("Al cerrar esta aplicación detendrá el servicio de envío de mensajes SMS a los agentes de venta,"
                + "\n ¿Está usted seguro que desea continuar con esta acción?", "¿Desea Salir?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (drClose == DialogResult.No)
                e.Cancel = true;
        }

        private void pruebaLocalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            localFiles();
        }

        private void auditoriaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openAuditoria();
        }
    }
}
