package dao;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import entities.Director;

public class DirectorDAO {

	public boolean ErrorThrown = false;
	public String ErrorMessage = "";

	public List<Director> readAll(Connection conn) {
		ErrorThrown = false;
		List<Director> result = new ArrayList<Director>();

		String sQuery = "SELECT director_id, director_name, email, cellphone, sms, id_empresa "
				+ "FROM dim_directors ORDER BY director_id";
		Statement stmt = null;
		ResultSet rs = null;
		try {
			stmt = conn.createStatement();
			rs = stmt.executeQuery(sQuery);

			while (rs.next()) {
				Director director = new Director();
				director.setId(rs.getLong(1));
				director.setName(rs.getString(2));
				director.setEmail(rs.getString(3));
				director.setCellPhone(rs.getString(4));
				director.setSms(rs.getBoolean(5));
				director.setIdEnterprise(rs.getLong(6));
				result.add(director);
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

	public boolean update(long id, Boolean bSMS, Connection conn) {
		ErrorThrown = false;
		PreparedStatement st = null;
		try {
			st = conn
					.prepareStatement("UPDATE dim_directors SET sms=? WHERE director_id = ?");
			st.setBoolean(1, bSMS);
			st.setLong(2, id);
			st.executeUpdate();
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
			return false;
		} finally {
			try {
				st.close();
			} catch (SQLException e) {
				// If too many statments or resultsets stay opened, the server
				// will have to be restarted.
			}
		}
		return true;
	}

	public boolean update(Director director, Connection conn) {
		ErrorThrown = false;
		PreparedStatement st = null;
		String strSQL = "UPDATE dim_directors "
				+ "SET director_name=?, email=?, cellphone=?, "
				+ "id_empresa=? " + "WHERE director_id=?";
		try {
			st = conn.prepareStatement(strSQL);
			st.setString(1, director.getName());
			st.setString(2, director.getEmail());
			st.setString(3, director.getCellPhone());
			st.setLong(4, director.getIdEnterprise());
			st.setLong(5, director.getId());
			st.executeUpdate();
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
			return false;
		}
		try {
			st.close();
		} catch (SQLException e) {
			// If too many statments or resultsets stay opened, the server will
			// have to be restarted.
		}
		return true;
	}

	public boolean delete(long id, Connection conn) {
		PreparedStatement st = null;
		try {
			st = conn
					.prepareStatement("DELETE FROM dim_directors WHERE director_id = ?");
			st.setLong(1, id);
			st.executeUpdate();
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
			return false;
		} finally {
			try {
				st.close();
			} catch (SQLException e) {
				// If too many statments or resultsets stay opened, the server
				// will have to be restarted.
			}
		}
		return true;
	}

	public Director create(Director director, Connection conn) {
		ErrorThrown = false;
		PreparedStatement st = null;
		String strSQL = "INSERT INTO dim_directors( "
				+ "director_name, email, cellphone, id_empresa) "
				+ "VALUES (?, ?, ?, ?)";
		try {
			st = conn.prepareStatement(strSQL, Statement.RETURN_GENERATED_KEYS);
			st.setString(1, director.getName());
			st.setString(2, director.getEmail());
			st.setString(3, director.getCellPhone());
			st.setLong(4, director.getIdEnterprise());
			int affectedRows = st.executeUpdate();
			if (affectedRows == 0) {
				ErrorThrown = true;
				ErrorMessage = "Error al intentar crear registro.";
				return null;
			}
			st.getGeneratedKeys().next();
			director.setId(st.getGeneratedKeys().getLong(1));
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
			return null;
		}
		try {
			st.close();
		} catch (SQLException e) {
			// If too many statments or resultsets stay opened, the server will
			// have to be restarted.
		}
		return director;
	}
}
