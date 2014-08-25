package dao;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class ConnectionManager {
	private Connection conn;
	 public Connection getConnection() throws SQLException, ClassNotFoundException {
		 if (conn== null)
			 Class.forName("org.postgresql.Driver");
			 conn = DriverManager.getConnection("jdbc:postgresql://localhost:5432/rhjasper", "rhjasper","rhjasper");
		 return conn;
	 }
}
