package dao;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import entities.SMS_Settings;

public class SMSSettingsDAO {

	public boolean ErrorThrown = false;
	public String ErrorMessage = "";

	public List<SMS_Settings> readAll(Connection conn) {
		ErrorThrown = false;
		List<SMS_Settings> result = new ArrayList<SMS_Settings>();

		String sQuery = "SELECT friexecuted, frischedule, monexecuted, monschedule, sms_password, "
				+ "satexecuted, satschedule, sunexecuted, sunschedule, thuexecuted, "
				+ "thuschedule, todschedule, tueexecuted, tueschedule, sms_user, "
				+ "wedexecuted, wedschedule " + "FROM sys_sms_settings";
		Statement stmt = null;
		ResultSet rs = null;
		try {
			stmt = conn.createStatement();
			rs = stmt.executeQuery(sQuery);

			while (rs.next()) {
				SMS_Settings SMS_Settings = new SMS_Settings();
				SMS_Settings.setFriExecuted(rs.getString(1));
				SMS_Settings.setFriSchedule(rs.getString(2));
				SMS_Settings.setMonExecuted(rs.getString(3));
				SMS_Settings.setMonSchedule(rs.getString(4));
				SMS_Settings.setSms_password(rs.getString(5));
				SMS_Settings.setSatExecuted(rs.getString(6));
				SMS_Settings.setSatSchedule(rs.getString(7));
				SMS_Settings.setSunExecuted(rs.getString(8));
				SMS_Settings.setSunSchedule(rs.getString(9));
				SMS_Settings.setThuExecuted(rs.getString(10));
				SMS_Settings.setThuSchedule(rs.getString(11));
				SMS_Settings.setTodSchedule(rs.getString(12));
				SMS_Settings.setTueExecuted(rs.getString(13));
				SMS_Settings.setTueSchedule(rs.getString(14));
				SMS_Settings.setSms_user(rs.getString(15));
				SMS_Settings.setWedExecuted(rs.getString(16));
				SMS_Settings.setWedSchedule(rs.getString(17));
				result.add(SMS_Settings);
			}
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
		}finally{
			try {
				stmt.close();
				rs.close();
			} catch (SQLException e) {
				//If too many statments or resultsets stay opened, the server will have to be restarted.
			}
		}
		return result;
	}

	public boolean update(SMS_Settings setting, Connection conn) {
		ErrorThrown = false;

		PreparedStatement deleteSettings = null;
		PreparedStatement insertSettings = null;

		String deleteSQL = "delete from sys_sms_settings";

		String insertSQL = "INSERT INTO sys_sms_settings("
				+ "friexecuted, frischedule, monexecuted, monschedule, sms_password,"
				+ "satexecuted, satschedule, sunexecuted, sunschedule, thuexecuted, "
				+ "thuschedule, todschedule, tueexecuted, tueschedule, sms_user, "
				+ "wedexecuted, wedschedule)" + "VALUES (?, ?, ?, ?, ?, "
				+ "?, ?, ?, ?, ?, " + "?, ?, ?, ?, ?, " + "?, ?);";

		try {
			conn.setAutoCommit(false);

			deleteSettings = conn.prepareStatement(deleteSQL);
			deleteSettings.executeUpdate();

			insertSettings = conn.prepareStatement(insertSQL);
			insertSettings.setString(1, setting.getFriExecuted());
			insertSettings.setString(2, setting.getFriSchedule());
			insertSettings.setString(3, setting.getMonExecuted());
			insertSettings.setString(4, setting.getMonSchedule());
			insertSettings.setString(5, setting.getSms_password());
			insertSettings.setString(6, setting.getSatExecuted());
			insertSettings.setString(7, setting.getSatSchedule());
			insertSettings.setString(8, setting.getSunExecuted());
			insertSettings.setString(9, setting.getSunSchedule());
			insertSettings.setString(10, setting.getThuExecuted());
			insertSettings.setString(11, setting.getThuSchedule());
			insertSettings.setString(12, setting.getTodSchedule());
			insertSettings.setString(13, setting.getTueExecuted());
			insertSettings.setString(14, setting.getTueSchedule());
			insertSettings.setString(15, setting.getSms_user());
			insertSettings.setString(16, setting.getWedExecuted());
			insertSettings.setString(17, setting.getWedSchedule());
			insertSettings.executeUpdate();

			conn.commit();
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
			return false;
		} finally {
			try {
				conn.setAutoCommit(true);
				if (deleteSettings != null) {
					deleteSettings.close();
				}
				if (insertSettings != null) {
					insertSettings.close();
				}
			} catch (SQLException e) {
				//If too many statments or resultsets stay opened, the server will have to be restarted.
			}
		}
		return true;
	}
}
