package dao;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import entities.Seller;

public class SellerDAO {

	public boolean ErrorThrown = false;
	public String ErrorMessage = "";

	public List<Seller> readAll(Connection conn) {
		ErrorThrown = false;
		List<Seller> result = new ArrayList<Seller>();

		String sQuery = "SELECT seller_id, sms, ap_id, agent_code, agent_name, email, cellphone, "
				+ "weekly_goal, id_empresa " + "FROM dim_sellers ORDER BY agent_name";

		Statement stmt = null;
		ResultSet rs = null;
		try {
			stmt = conn.createStatement();
			rs = stmt.executeQuery(sQuery);

			while (rs.next()) {
				Seller seller = new Seller();
				seller.setId(rs.getLong(1));
				seller.setSms(rs.getBoolean(2));
				seller.setAp_id(rs.getLong(3));
				seller.setCode(rs.getString(4));
				seller.setName(rs.getString(5));
				seller.setEmail(rs.getString(6));
				seller.setCellPhone(rs.getString(7));
				seller.setWeeklyGoal(rs.getFloat(8));
				seller.setIdEnterprise(rs.getLong(9));
				result.add(seller);
			}
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
		}
		finally{
			try {
				stmt.close();
				rs.close();
			} catch (SQLException e) {
				// If too many statments or resultsets stay opened, the server will have to be restarted. 
			}
		}
		return result;
	}

	public boolean update(long idSeller, boolean bSMS, Connection conn) {
		ErrorThrown = false;
		PreparedStatement st = null;
		String strSQL="UPDATE dim_sellers " +
				"SET sms=? WHERE seller_id=?";
		try {
			st = conn.prepareStatement(strSQL);
			st.setBoolean(1, bSMS);
			st.setLong(2, idSeller);
			st.executeUpdate();
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
			return false;
		}try {
			st.close();
		} catch (SQLException e) {
			// If too many statments or resultsets stay opened, the server will have to be restarted. 
		}
		return true;
	}
	public boolean update(Seller seller, Connection conn) {
		ErrorThrown = false;
		PreparedStatement st = null;
		String strSQL="UPDATE dim_sellers " +
				"SET agent_name=?, email=?, cellphone=?, weekly_goal=? " +
				"WHERE seller_id=?";
		try {
			st = conn.prepareStatement(strSQL);
			st.setString(1, seller.getName());
			st.setString(2, seller.getEmail());
			st.setString(3, seller.getCellPhone());
			st.setFloat(4, seller.getWeeklyGoal());
			st.setLong(5, seller.getId());
			st.executeUpdate();
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
			return false;
		}try {
			st.close();
		} catch (SQLException e) {
			// If too many statments or resultsets stay opened, the server will have to be restarted. 
		}
		return true;
	}
	public Seller create(Seller seller, Connection conn){		
		ErrorThrown = false;
		PreparedStatement st = null;
		String strSQL= "INSERT INTO dim_sellers( " +
				"agent_name, email, cellphone, " + 
				"weekly_goal, id_empresa, ap_id) " +
				"VALUES (?, ?, ?, ?, ?, ?)";
		try {
			st = conn.prepareStatement(strSQL, Statement.RETURN_GENERATED_KEYS);
			st.setString(1, seller.getName());
			st.setString(2, seller.getEmail());
			st.setString(3, seller.getCellPhone());
			st.setFloat(4, seller.getWeeklyGoal());
			st.setLong(5, seller.getIdEnterprise());
			st.setLong(6, seller.getAp_id());
			int affectedRows = st.executeUpdate();
			if(affectedRows == 0) {
				ErrorThrown =true;
				ErrorMessage = "Error al intentar crear registro.";
				return null;
			}
			st.getGeneratedKeys().next();
			seller.setId(st.getGeneratedKeys().getLong(1));
		} catch (SQLException e) {
			ErrorThrown = true;
			ErrorMessage = e.getMessage();
			return null;
		}try {
			st.close();
		} catch (SQLException e) {
			// If too many statments or resultsets stay opened, the server will have to be restarted. 
		}
		return seller;
	}
}
