package gateway;

import java.io.IOException;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import javax.naming.Context;
import javax.naming.InitialContext;
import javax.naming.NamingException;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.sql.DataSource;

import com.google.gson.Gson;

import dao.DirectorDAO;
import dao.EnterpriseDAO;
import dao.SMSSettingsDAO;
import dao.SellerDAO;
import entities.Director;
import entities.Enterprise;
import entities.SMS_Settings;
import entities.Seller;

/**
 * Servlet implementation class GatewayServlet
 */
@WebServlet(name = "/GatewayServlet", urlPatterns = { "/GatewayServlet" })
public class GatewayServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;

	/**
	 * @see HttpServlet#HttpServlet()
	 */
	public GatewayServlet() {
		super();
		// TODO Auto-generated constructor stub
	}

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse
	 *      response)
	 */
	protected void doGet(HttpServletRequest request,
			HttpServletResponse response) throws ServletException, IOException {

	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse
	 *      response)
	 */
	protected void doPost(HttpServletRequest request,
			HttpServletResponse response) throws ServletException, IOException {
		GatewayResponse baseResponse = new GatewayResponse();
		PrintWriter out = response.getWriter();
		Gson gson = new Gson();
		Context ctx;
		Connection conn;

		try {
			ctx = new InitialContext();
			DataSource ds = (DataSource) ctx
					.lookup("java:/comp/env/jdbc/smsPostgres");
			conn = ds.getConnection();
			
			SellerDAO sellerDAO = new SellerDAO();
			DirectorDAO directorDAO = new DirectorDAO();
			SMSSettingsDAO settingsDAO = new SMSSettingsDAO();
			List<SMS_Settings> settingsList;
			String sRequest;
			
			String command = request.getParameter("cmd");
			switch (command) {
			case "readSellers":

				List<Seller> sellerList;

				sellerList = sellerDAO.readAll(conn);
				if (sellerDAO.ErrorThrown) {
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(sellerDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}

				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("OK");
				baseResponse.setResult(sellerList);
				out.println(gson.toJson(baseResponse));
				break;
			case "readDirectors":

				List<Director> directorList;

				directorList = directorDAO.readAll(conn);
				if (directorDAO.ErrorThrown) {
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(directorDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}

				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("OK");
				baseResponse.setResult(directorList);
				out.println(gson.toJson(baseResponse));
				break;
			case "readEnterprises":

				List<Enterprise> enterpriseList;
				EnterpriseDAO enterpriseDAO = new EnterpriseDAO();
				enterpriseList = enterpriseDAO.readAll(conn);
				if (enterpriseDAO.ErrorThrown) {
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(enterpriseDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}

				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("OK");
				baseResponse.setResult(enterpriseList);
				out.println(gson.toJson(baseResponse));
				break;
			case "readSettings":

				settingsList = settingsDAO.readAll(conn);
				if (settingsDAO.ErrorThrown) {
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(settingsDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}

				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("OK");
				baseResponse.setResult(settingsList);
				out.println(gson.toJson(baseResponse));
				break;
			case "sellerSMSUpdate":
				String idSeller = request.getParameter("sellerid");
				String bSMS = request.getParameter("sms");

				if (!sellerDAO.update(Long.parseLong(idSeller),
						Boolean.parseBoolean(bSMS), conn)
						|| sellerDAO.ErrorThrown) {
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(sellerDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}

				baseResponse.setErrorThrown(false);
				baseResponse
						.setResponseDescription("Propiedad SMS actualizada exitosamente!");
				out.println(gson.toJson(baseResponse));
				break;
			case "directorSMSUpdate":
				String idDirector = request.getParameter("directorid");
				String bSMSDirector = request.getParameter("sms");

				if (!directorDAO.update(Long.parseLong(idDirector),
						Boolean.parseBoolean(bSMSDirector), conn)
						|| directorDAO.ErrorThrown) {
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(directorDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}

				baseResponse.setErrorThrown(false);
				baseResponse
						.setResponseDescription("Propiedad SMS actualizada exitosamente!");
				out.println(gson.toJson(baseResponse));
				break;
			case "settingsUpdate":
				StringBuilder sb = new StringBuilder();
				while ((sRequest = request.getReader().readLine()) != null) {
					sb.append(sRequest);
				}
				SMS_Settings setting = (SMS_Settings) gson.fromJson(sb.toString(),
						SMS_Settings.class);
				if(!settingsDAO.update(setting, conn)){
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(settingsDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}
				settingsList = new ArrayList<SMS_Settings>();
				settingsList.add(setting);
				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("Configuracion actualizada exitosamente!");
				baseResponse.setResult(setting);
				out.println(gson.toJson(baseResponse));
				break;
			case "sellerUpdate":
				StringBuilder sbUpdate = new StringBuilder();
				
				while ((sRequest = request.getReader().readLine()) != null) {
					sbUpdate.append(sRequest);
				}
				Seller sellerToUpdate = (Seller) gson.fromJson(sbUpdate.toString(),
						Seller.class);
				if(!sellerDAO.update(sellerToUpdate, conn)){
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(sellerDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}
				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("Agente actualizado exitosamente!");
				baseResponse.setResult("OK");
				out.println(gson.toJson(baseResponse));
				break;
			case "sellerCreate":
				StringBuilder sbSellerCreate = new StringBuilder();
				
				while ((sRequest = request.getReader().readLine()) != null) {
					sbSellerCreate.append(sRequest);
				}
				Seller sellerToCreate = (Seller) gson.fromJson(sbSellerCreate.toString(),
						Seller.class);
				sellerToCreate.setAp_id(-1);
				Seller result = sellerDAO.create(sellerToCreate, conn);
				if(sellerDAO.ErrorThrown){
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(sellerDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}
				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("Agente creado exitosamente!");
				baseResponse.setResult(result);
				out.println(gson.toJson(baseResponse));
				break;
			case "directorUpdate":
				StringBuilder sbUpdateDirector = new StringBuilder();
				
				while ((sRequest = request.getReader().readLine()) != null) {
					sbUpdateDirector.append(sRequest);
				}
				Director directorToUpdate = (Director) gson.fromJson(sbUpdateDirector.toString(),
						Director.class);
				if(!directorDAO.update(directorToUpdate, conn)){
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(directorDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}
				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("Director actualizado exitosamente!");
				baseResponse.setResult("OK");
				out.println(gson.toJson(baseResponse));
				break;
			case "directorCreate":
				StringBuilder sbDirectorCreate = new StringBuilder();
				
				while ((sRequest = request.getReader().readLine()) != null) {
					sbDirectorCreate.append(sRequest);
				}
				Director directorToCreate = (Director) gson.fromJson(sbDirectorCreate.toString(),
						Director.class);
				Director resultDirector = directorDAO.create(directorToCreate, conn);
				if(directorDAO.ErrorThrown){
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(directorDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}
				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("Director creado exitosamente!");
				baseResponse.setResult(resultDirector);
				out.println(gson.toJson(baseResponse));
				break;
			case "directorDelete":
				String idDirectorToDelete = request.getParameter("directorid");

				if (!directorDAO.delete(Long.parseLong(idDirectorToDelete), conn)
						|| directorDAO.ErrorThrown) {
					baseResponse.setErrorThrown(true);
					baseResponse.setResponseDescription(directorDAO.ErrorMessage);
					out.println(gson.toJson(baseResponse));
					break;
				}

				baseResponse.setErrorThrown(false);
				baseResponse.setResponseDescription("Registro de Director eliminado exitosamente.");
				out.println(gson.toJson(baseResponse));
				break;
			}
			conn.close();
			
		} catch (NamingException | SQLException e1) {
			baseResponse.setErrorThrown(true);
			baseResponse.setResponseDescription(e1.getMessage());
			out.println(gson.toJson(baseResponse));
		}
	}
}
