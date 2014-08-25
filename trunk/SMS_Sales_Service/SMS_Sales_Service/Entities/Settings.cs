
using Npgsql;
using SMSSender;
using System;
namespace SMS_Sales_Service.Entities
{
    public class Settings
    {
        public string friExecuted = "";
        public string friSchedule = "";
        public string monExecuted = "";
        public string monSchedule = "";
        public string sms_password = "";
        public string satExecuted = "";
        public string satSchedule = "";
        public string sunExecuted = "";
        public string sunSchedule = "";
        public string thuExecuted = "";
        public string thuSchedule = "";
        public string todSchedule = "";
        public string tueExecuted = "";
        public string tueSchedule = "";
        public string sms_user = "";
        public string wedExecuted = "";
        public string wedSchedule = "";

        public string ToString(){
            return "monSchedule: " + monSchedule + Environment.NewLine +
                "tueSchedule: " + tueSchedule + Environment.NewLine +
                "wedSchedule: " + wedSchedule + Environment.NewLine +
                "thuSchedule: " + thuSchedule + Environment.NewLine +
                "friSchedule: " + friSchedule + Environment.NewLine +
                "satSchedule: " + satSchedule + Environment.NewLine +
                "sunSchedule: " + sunSchedule + Environment.NewLine +
                "todSchedule: " + todSchedule + Environment.NewLine +
                "sms_username: " + sms_user;
        }
        
        public void Save(System.Diagnostics.EventLog eventLog)
        {
            ConnectionManager connectionManager = new ConnectionManager();
            NpgsqlConnection conn = connectionManager.getConnection();
            string strInsert = "INSERT INTO sys_sms_settings( " +
            "friexecuted, frischedule, monexecuted, monschedule, sms_password, " +
            "satexecuted, satschedule, sunexecuted, sunschedule, thuexecuted, " +
            "thuschedule, todschedule, tueexecuted, tueschedule, sms_user, " +
            "wedexecuted, wedschedule) " +
            "VALUES ('" + friExecuted + "', '" + friSchedule + "', '" + monExecuted + "', '" + monSchedule + "', '" + sms_password + "', '" +
            satExecuted + "', '" + satSchedule + "', '" + sunExecuted + "', '" + sunSchedule + "', '" + thuExecuted + "', '" +
            thuSchedule + "', '" + todSchedule + "', '" + tueExecuted + "', '" + tueSchedule + "', '" + sms_user + "', '" +
            wedExecuted + "', '" + wedSchedule + "')";

            NpgsqlCommand command = new NpgsqlCommand(strInsert, conn);
            int rowsaffected;

            try
            {
                new NpgsqlCommand("delete from sys_sms_settings", conn).ExecuteNonQuery();
                rowsaffected = command.ExecuteNonQuery();
                eventLog.WriteEntry("Settings update done.", System.Diagnostics.EventLogEntryType.Information);
            }
            finally
            {
                conn.Close();
            }
        }
    }
}
