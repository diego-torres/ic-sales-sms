using System.Collections.Generic;
using Npgsql;
using System.Data;
using SMS_Sales_Service.Entities;

namespace SMSSender.DAO
{
    public class SettingsDAO
    {
        public bool ErrorThrown = false;
        public string ErrorMessage = "";
        public Settings settings = new Settings();
        public Settings readAll(NpgsqlConnection conn)
        {
            Settings setting = new Settings();

            string sql = "SELECT friexecuted, frischedule, monexecuted, monschedule, sms_password, " +
                            "satexecuted, satschedule, sunexecuted, sunschedule, thuexecuted, " +
                            "thuschedule, todschedule, tueexecuted, tueschedule, sms_user, " +
                            "wedexecuted, wedschedule FROM sys_sms_settings";

            DataTable dt = new DataTable();
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);

            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                setting.friExecuted = dt.Rows[0][0].ToString();
                setting.friSchedule = dt.Rows[0][1].ToString();
                setting.monExecuted = dt.Rows[0][2].ToString();
                setting.monSchedule = dt.Rows[0][3].ToString();
                setting.sms_password = dt.Rows[0][4].ToString();
                setting.satExecuted = dt.Rows[0][5].ToString();
                setting.satSchedule = dt.Rows[0][6].ToString();
                setting.sunExecuted = dt.Rows[0][7].ToString();
                setting.sunSchedule = dt.Rows[0][8].ToString();
                setting.thuExecuted = dt.Rows[0][9].ToString();
                setting.thuSchedule = dt.Rows[0][10].ToString();
                setting.todSchedule = dt.Rows[0][11].ToString();
                setting.tueExecuted = dt.Rows[0][12].ToString();
                setting.tueSchedule = dt.Rows[0][13].ToString();
                setting.sms_user = dt.Rows[0][14].ToString();
                setting.wedExecuted = dt.Rows[0][15].ToString();
                setting.wedSchedule = dt.Rows[0][16].ToString();

            }

            da.Dispose();

            this.settings = setting;
            return setting;
        }
    }
}
