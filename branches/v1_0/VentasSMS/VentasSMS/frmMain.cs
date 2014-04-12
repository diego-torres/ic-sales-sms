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

namespace VentasSMS
{
    public partial class frmMain : Form
    {
        private AdminPaqImpl api;
        public AdminPaqImpl API { get { return api; } }
        private frmConfig fConfig = null;
        private ScheduleConfig sConfig = null;

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

        private string configuredFilePath()
        {
            string configFilePath = ConfigurationManager.AppSettings[CommonConstants.SMS_FILE_NAME];
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
            DateTime today = DateTime.Today;
            string executionBackup="";
            switch (today.DayOfWeek)
            { 
                case DayOfWeek.Monday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.MON_EXECUTED];
                    break;
                case DayOfWeek.Tuesday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.TUE_EXECUTED];
                    break;
                case DayOfWeek.Wednesday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.WED_EXECUTED];
                    break;
                case DayOfWeek.Thursday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.THU_EXECUTED];
                    break;
                case DayOfWeek.Friday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.FRI_EXECUTED];
                    break;
                case DayOfWeek.Saturday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.SAT_EXECUTED];
                    break;
                case DayOfWeek.Sunday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.SUN_EXECUTED];
                    break;
            }

            // Open App.Config of executable
            System.Configuration.Configuration config =
             ConfigurationManager.OpenExeConfiguration
                        (ConfigurationUserLevel.None);

            config.AppSettings.Settings.Remove(CommonConstants.MON_EXECUTED);
            config.AppSettings.Settings.Remove(CommonConstants.TUE_EXECUTED);
            config.AppSettings.Settings.Remove(CommonConstants.WED_EXECUTED);
            config.AppSettings.Settings.Remove(CommonConstants.THU_EXECUTED);
            config.AppSettings.Settings.Remove(CommonConstants.FRI_EXECUTED);
            config.AppSettings.Settings.Remove(CommonConstants.SAT_EXECUTED);
            config.AppSettings.Settings.Remove(CommonConstants.SUN_EXECUTED);


            config.AppSettings.Settings.Add(CommonConstants.MON_EXECUTED, today.DayOfWeek==DayOfWeek.Monday ? executionBackup:"");
            config.AppSettings.Settings.Add(CommonConstants.TUE_EXECUTED, today.DayOfWeek == DayOfWeek.Tuesday ? executionBackup : "");
            config.AppSettings.Settings.Add(CommonConstants.WED_EXECUTED, today.DayOfWeek == DayOfWeek.Wednesday ? executionBackup : "");
            config.AppSettings.Settings.Add(CommonConstants.THU_EXECUTED, today.DayOfWeek == DayOfWeek.Thursday ? executionBackup : "");
            config.AppSettings.Settings.Add(CommonConstants.FRI_EXECUTED, today.DayOfWeek == DayOfWeek.Friday ? executionBackup : "");
            config.AppSettings.Settings.Add(CommonConstants.SAT_EXECUTED, today.DayOfWeek == DayOfWeek.Saturday ? executionBackup : "");
            config.AppSettings.Settings.Add(CommonConstants.SUN_EXECUTED, today.DayOfWeek == DayOfWeek.Sunday ? executionBackup : "");

            // Save the changes in App.config file.
            config.Save(ConfigurationSaveMode.Modified);

            // Force a reload of a changed section.
            ConfigurationManager.RefreshSection("appSettings");

        }

        private bool HourExecuted (string current)
        {
            DateTime today = DateTime.Today;
            string sExecuted = "";
            string[] executed;
            int found = 0;

            switch (today.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    sExecuted = ConfigurationManager.AppSettings[CommonConstants.MON_EXECUTED];
                    break;
                case DayOfWeek.Tuesday:
                    sExecuted = ConfigurationManager.AppSettings[CommonConstants.TUE_EXECUTED];
                    break;
                case DayOfWeek.Wednesday:
                    sExecuted = ConfigurationManager.AppSettings[CommonConstants.WED_EXECUTED];
                    break;
                case DayOfWeek.Thursday:
                    sExecuted = ConfigurationManager.AppSettings[CommonConstants.THU_EXECUTED];
                    break;
                case DayOfWeek.Friday:
                    sExecuted = ConfigurationManager.AppSettings[CommonConstants.FRI_EXECUTED];
                    break;
                case DayOfWeek.Saturday:
                    sExecuted = ConfigurationManager.AppSettings[CommonConstants.SAT_EXECUTED];
                    break;
                case DayOfWeek.Sunday:
                    sExecuted = ConfigurationManager.AppSettings[CommonConstants.SUN_EXECUTED];
                    break;
            }
            executed = sExecuted.Split(';');
            found = Array.IndexOf(executed, current);
            return found>=0;
        }

        private void addExecutionToList(string sTime)
        {
            DateTime today = DateTime.Today;
            string executionBackup = "";
            // Open App.Config of executable
            System.Configuration.Configuration config =
             ConfigurationManager.OpenExeConfiguration
                        (ConfigurationUserLevel.None);
            switch (today.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.MON_EXECUTED];
                    config.AppSettings.Settings.Remove(CommonConstants.MON_EXECUTED);
                    config.AppSettings.Settings.Add(CommonConstants.MON_EXECUTED, executionBackup.Equals("") ? sTime: executionBackup + ";" + sTime);
                    break;
                case DayOfWeek.Tuesday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.TUE_EXECUTED];
                    config.AppSettings.Settings.Remove(CommonConstants.TUE_EXECUTED);
                    config.AppSettings.Settings.Add(CommonConstants.TUE_EXECUTED, executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime);
                    break;
                case DayOfWeek.Wednesday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.WED_EXECUTED];
                    config.AppSettings.Settings.Remove(CommonConstants.WED_EXECUTED);
                    config.AppSettings.Settings.Add(CommonConstants.WED_EXECUTED, executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime);
                    break;
                case DayOfWeek.Thursday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.THU_EXECUTED];
                    config.AppSettings.Settings.Remove(CommonConstants.THU_EXECUTED);
                    config.AppSettings.Settings.Add(CommonConstants.THU_EXECUTED, executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime);
                    break;
                case DayOfWeek.Friday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.FRI_EXECUTED];
                    config.AppSettings.Settings.Remove(CommonConstants.FRI_EXECUTED);
                    config.AppSettings.Settings.Add(CommonConstants.FRI_EXECUTED, executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime);
                    break;
                case DayOfWeek.Saturday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.SAT_EXECUTED];
                    config.AppSettings.Settings.Remove(CommonConstants.SAT_EXECUTED);
                    config.AppSettings.Settings.Add(CommonConstants.SAT_EXECUTED, executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime);
                    break;
                case DayOfWeek.Sunday:
                    executionBackup = ConfigurationManager.AppSettings[CommonConstants.SUN_EXECUTED];
                    config.AppSettings.Settings.Remove(CommonConstants.SUN_EXECUTED);
                    config.AppSettings.Settings.Add(CommonConstants.SUN_EXECUTED, executionBackup.Equals("") ? sTime : executionBackup + ";" + sTime);
                    break;
            }

            // Save the changes in App.config file.
            config.Save(ConfigurationSaveMode.Modified);

            // Force a reload of a changed section.
            ConfigurationManager.RefreshSection("appSettings");

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

        private void sendSMS()
        {
            string configFilePath = configuredFilePath();

            if ("".Equals(configFilePath.Trim()) || !File.Exists(configFilePath))
            {
                openConfigForm(configFilePath);
            }
            else
            {
                SalesPicker sp = new SalesPicker();
                sp.WorkbookPath = configFilePath;
                sp.UserName = ConfigurationManager.AppSettings[CommonConstants.USER];
                sp.Password = ConfigurationManager.AppSettings[CommonConstants.PASSWORD];
                //sp.SendMethod = SMSSender.CommonConstants.SMS_LOCAL;
                sp.SendMethod = SMSSender.CommonConstants.SMS_MAS_MENSAJES;
                
                sp.RetrieveConfiguration();
                sp.RetrieveSales();
                sp.SendSMS();
                sp.SendBossSMS();
            }
        }

        private bool DayEnabled()
        {
            DateTime today = DateTime.Today;
            
            switch (today.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    return bool.Parse(ConfigurationManager.AppSettings[CommonConstants.MON_ENABLED]);
                case DayOfWeek.Tuesday:
                    return bool.Parse(ConfigurationManager.AppSettings[CommonConstants.TUE_ENABLED]);
                case DayOfWeek.Wednesday:
                    return bool.Parse(ConfigurationManager.AppSettings[CommonConstants.WED_ENABLED]);
                case DayOfWeek.Thursday:
                    return bool.Parse(ConfigurationManager.AppSettings[CommonConstants.THU_ENABLED]);
                case DayOfWeek.Friday:
                    return bool.Parse(ConfigurationManager.AppSettings[CommonConstants.FRI_ENABLED]);
                case DayOfWeek.Saturday:
                    return bool.Parse(ConfigurationManager.AppSettings[CommonConstants.SAT_ENABLED]);
                case DayOfWeek.Sunday:
                    return bool.Parse(ConfigurationManager.AppSettings[CommonConstants.SUN_ENABLED]);
            }
            return false;
        }

        private void timerSMS_Tick(object sender, EventArgs e)
        {
            timerSMS.Stop();
            if (!DayEnabled()) 
            {
                timerSMS.Start();
                return;
            }
            

            resetWeek();
            string sScheduled = ConfigurationManager.AppSettings[CommonConstants.TOD_SCHEDULE];
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
    }
}
