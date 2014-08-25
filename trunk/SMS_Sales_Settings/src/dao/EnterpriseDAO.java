package dao;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import entities.Enterprise;

public class EnterpriseDAO {

	public boolean ErrorThrown = false;
	public String ErrorMessage = "";

	public List<Enterprise> readAll(Connection conn) {
		ErrorThrown = false;
		List<Enterprise> result = new ArrayList<Enterprise>();

		String sQuery = "SELECT distinct id_empresa, empresa FROM dim_sellers where id_empresa > -1";
		Statement stmt = null;
		ResultSet rs = null;
		try {
			stmt = conn.createStatement();
			rs = stmt.executeQuery(sQuery);

			while (rs.next()) {
				Enterprise enterprise = new Enterprise();
				enterprise.setId(rs.getLong(1));
				enterprise.setName(rs.getString(2));
				result.add(enterprise);
			}
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
		} finally {
			try {
				stmt.close();
				rs.close();
			} catch (SQLException e) {
				// If too many statments or resultsets stay opened, the server
				// will have to be restarted.
			}
		}
		return result;
	}
}
