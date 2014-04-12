using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace VentasSMS
{
    public partial class ScheduleConfig : Form
    {
        private Schedule schedule;
        public Schedule ExecutionSchedule { get { return schedule; } set { schedule = value; } }
        private bool dirty = false;

        public ScheduleConfig()
        {
            InitializeComponent();
        }


        private Schedule configuredSchedule()
        {
            Schedule result = new Schedule();
            string sTodSchedule, sMonEnabled, sTueEnabled, sWedEnabled, sThuEnabled, sFriEnabled, 
                sSatEnabled, sSunEnabled, user, password;

            sTodSchedule = ConfigurationManager.AppSettings[CommonConstants.TOD_SCHEDULE];

            result.Horarios = sTodSchedule.Split(';');
            Array.Sort(result.Horarios);

            sMonEnabled = ConfigurationManager.AppSettings[CommonConstants.MON_ENABLED];
            result.IsMonday = bool.Parse(sMonEnabled);

            sTueEnabled = ConfigurationManager.AppSettings[CommonConstants.TUE_ENABLED];
            result.IsTuesday = bool.Parse(sTueEnabled);

            sWedEnabled = ConfigurationManager.AppSettings[CommonConstants.WED_ENABLED];
            result.IsWednesday = bool.Parse(sWedEnabled);

            sThuEnabled = ConfigurationManager.AppSettings[CommonConstants.THU_ENABLED];
            result.IsThursday = bool.Parse(sThuEnabled);

            sFriEnabled = ConfigurationManager.AppSettings[CommonConstants.FRI_ENABLED];
            result.IsFriday = bool.Parse(sFriEnabled);

            sSatEnabled = ConfigurationManager.AppSettings[CommonConstants.SAT_ENABLED];
            result.IsSaturday = bool.Parse(sSatEnabled);

            sSunEnabled = ConfigurationManager.AppSettings[CommonConstants.SUN_ENABLED];
            result.IsSunday = bool.Parse(sSunEnabled);

            user = ConfigurationManager.AppSettings[CommonConstants.USER];
            password = ConfigurationManager.AppSettings[CommonConstants.PASSWORD];

            result.User = user;
            result.Password = password;

            return result;
        }

        public void LoadConfiguredSchedule()
        {
            this.schedule = configuredSchedule();
        }

        private void setConfiguredSchedule()
        {
            this.checkBoxMon.Checked = schedule.IsMonday;
            this.checkBoxTue.Checked = schedule.IsTuesday;
            this.checkBoxWed.Checked = schedule.IsWednesday;
            this.checkBoxThu.Checked = schedule.IsThursday;
            this.checkBoxFri.Checked = schedule.IsFriday;

            this.checkBoxSat.Checked = schedule.IsSaturday;
            this.checkBoxSun.Checked = schedule.IsSunday;

            this.listBoxSchedule.Items.Clear();

            foreach (string hour in schedule.Horarios)
            {
                this.listBoxSchedule.Items.Add(hour);
            }

            this.textBoxUser.Text = schedule.User;
            this.textBoxPassword.Text = schedule.Password;
        }

        private void ScheduleConfig_Load(object sender, EventArgs e)
        {
            LoadConfiguredSchedule();
            setConfiguredSchedule();
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            string selectedHour = this.dateTimePicker.Value.ToString("HH:mm");
            if (Array.IndexOf(schedule.Horarios, selectedHour) >= 0)
            {
                MessageBox.Show("El horario seleccionado ya ha sido programado previamente", "horario duplicado!", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return;
            }
            string[] horarios = schedule.Horarios;

            Array.Resize<string>(ref horarios, schedule.Horarios.Length + 1);
            horarios[schedule.Horarios.Length] = selectedHour;
            Array.Sort(horarios);
            schedule.Horarios = horarios;
            this.listBoxSchedule.Items.Clear();
            foreach (string hour in schedule.Horarios)
            {
                this.listBoxSchedule.Items.Add(hour);
            }
            dirty = true;
        }

        private void buttonRemove_Click(object sender, EventArgs e)
        {
            if (this.listBoxSchedule.SelectedItem != null)
            {
                if (this.listBoxSchedule.Items.Count == 1)
                {
                    MessageBox.Show("Por lo menos debe dejar un horario en el cual ejecutar los mensajes", "Imposible eliminar último horario", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    return;
                }

                this.listBoxSchedule.Items.Remove(this.listBoxSchedule.SelectedItem);
                string[] newSchedule = new string[1];

                foreach (string s in listBoxSchedule.Items)
                {
                    Array.Resize<string>(ref newSchedule, newSchedule.Length + 1);
                    newSchedule[newSchedule.Length - 2] = s;
                }

                Array.Resize<string>(ref newSchedule, newSchedule.Length - 1);
                Array.Sort(newSchedule);
                schedule.Horarios = newSchedule;

                dirty = true;
            }
            else
            {
                MessageBox.Show("Seleccione un horario de la lista para eliminar", "Horario no seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Hand);
            }
        }

        private void checkBox_Click(object sender, EventArgs e)
        {
            dirty = true;
        }

        private void saveSettings()
        {
            // Open App.Config of executable
            System.Configuration.Configuration config =
             ConfigurationManager.OpenExeConfiguration
                        (ConfigurationUserLevel.None);

            config.AppSettings.Settings.Remove(CommonConstants.TOD_SCHEDULE);

            config.AppSettings.Settings.Remove(CommonConstants.MON_ENABLED);
            config.AppSettings.Settings.Remove(CommonConstants.TUE_ENABLED);
            config.AppSettings.Settings.Remove(CommonConstants.WED_ENABLED);
            config.AppSettings.Settings.Remove(CommonConstants.THU_ENABLED);
            config.AppSettings.Settings.Remove(CommonConstants.FRI_ENABLED);

            config.AppSettings.Settings.Remove(CommonConstants.SAT_ENABLED);
            config.AppSettings.Settings.Remove(CommonConstants.SUN_ENABLED);

            config.AppSettings.Settings.Remove(CommonConstants.USER);
            config.AppSettings.Settings.Remove(CommonConstants.PASSWORD);
            
            // Add an Application Setting.
            string todSchedule = "";
            foreach(string tod in schedule.Horarios)
            {
                todSchedule += tod + ";";
            }
            todSchedule = todSchedule.Substring(0, todSchedule.Length - 1);

            config.AppSettings.Settings.Add(CommonConstants.TOD_SCHEDULE, todSchedule);

            config.AppSettings.Settings.Add(CommonConstants.MON_ENABLED, this.checkBoxMon.Checked.ToString());
            config.AppSettings.Settings.Add(CommonConstants.TUE_ENABLED, this.checkBoxTue.Checked.ToString());
            config.AppSettings.Settings.Add(CommonConstants.WED_ENABLED, this.checkBoxWed.Checked.ToString());
            config.AppSettings.Settings.Add(CommonConstants.THU_ENABLED, this.checkBoxThu.Checked.ToString());
            config.AppSettings.Settings.Add(CommonConstants.FRI_ENABLED, this.checkBoxFri.Checked.ToString());

            config.AppSettings.Settings.Add(CommonConstants.SAT_ENABLED, this.checkBoxSat.Checked.ToString());
            config.AppSettings.Settings.Add(CommonConstants.SUN_ENABLED, this.checkBoxSun.Checked.ToString());

            config.AppSettings.Settings.Add(CommonConstants.USER, this.textBoxUser.Text);
            config.AppSettings.Settings.Add(CommonConstants.PASSWORD, this.textBoxPassword.Text);

            // Save the changes in App.config file.
            config.Save(ConfigurationSaveMode.Modified);

            // Force a reload of a changed section.
            ConfigurationManager.RefreshSection("appSettings");
        }

        private void ScheduleConfig_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dirty)
            {
                DialogResult save = MessageBox.Show("¿Desea guardar los cambios realizados a la agenda?", "¿Guardar cambios?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (save == DialogResult.Yes)
                {
                    saveSettings();
                    dirty = false;
                }
                else if (save == DialogResult.No)
                {
                    dirty = false;
                }
                else 
                {
                    e.Cancel = true;
                }
            }
        }

        private void toolStripButtonSave_Click(object sender, EventArgs e)
        {
            saveSettings();
            dirty = false;
        }

        private void toolStripButtonUndo_Click(object sender, EventArgs e)
        {
            LoadConfiguredSchedule();
            setConfiguredSchedule();
            dirty = false;
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        {
            schedule.User = textBoxUser.Text;
            schedule.Password = textBoxPassword.Text;
            dirty = true;
        }
    }

    public class Schedule
    {
        private string[] horarios;
        private bool mon, tue, wed, thu, fri, sat, sun;
        private string user, password;

        public string[] Horarios { get { return horarios; } set { horarios = value; } }
        public bool IsMonday { get { return mon; } set { mon = value; } }
        public bool IsTuesday { get { return tue; } set { tue = value; } }
        public bool IsWednesday { get { return wed; } set { wed = value; } }
        public bool IsThursday { get { return thu; } set { thu = value; } }
        public bool IsFriday { get { return fri; } set { fri = value; } }
        public bool IsSaturday { get { return sat; } set { sat = value; } }
        public bool IsSunday { get { return sun; } set { sun = value; } }
        public string User { get { return user; } set { user = value; } }
        public string Password { get { return password; } set { password = value; } }
    }
}
