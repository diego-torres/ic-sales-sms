package entities;

public class Seller {
	private long id;
    private boolean sms = false;
    private long ap_id = -1;
    private String code = "";
    private String name;
    private String email;
    private String cellPhone;
    private float weeklyGoal = 0;
    private float weeklySold = 0;
    private float monthlySold = 0;
    private long idEnterprise = -1;

    private Enterprise enterprise = null;

	/**
	 * @return the id
	 */
	public long getId() {
		return id;
	}

	/**
	 * @param id the id to set
	 */
	public void setId(long id) {
		this.id = id;
	}

	/**
	 * @return the sms
	 */
	public boolean isSms() {
		return sms;
	}

	/**
	 * @param sms the sms to set
	 */
	public void setSms(boolean sms) {
		this.sms = sms;
	}

	/**
	 * @return the ap_id
	 */
	public long getAp_id() {
		return ap_id;
	}

	/**
	 * @param ap_id the ap_id to set
	 */
	public void setAp_id(long ap_id) {
		this.ap_id = ap_id;
	}

	/**
	 * @return the code
	 */
	public String getCode() {
		return code;
	}

	/**
	 * @param code the code to set
	 */
	public void setCode(String code) {
		this.code = code;
	}

	/**
	 * @return the name
	 */
	public String getName() {
		return name;
	}

	/**
	 * @param name the name to set
	 */
	public void setName(String name) {
		this.name = name;
	}

	/**
	 * @return the email
	 */
	public String getEmail() {
		return email;
	}

	/**
	 * @param email the email to set
	 */
	public void setEmail(String email) {
		this.email = email;
	}

	/**
	 * @return the cellPhone
	 */
	public String getCellPhone() {
		return cellPhone;
	}

	/**
	 * @param cellPhone the cellPhone to set
	 */
	public void setCellPhone(String cellPhone) {
		this.cellPhone = cellPhone;
	}

	/**
	 * @return the weeklyGoal
	 */
	public float getWeeklyGoal() {
		return weeklyGoal;
	}

	/**
	 * @param weeklyGoal the weeklyGoal to set
	 */
	public void setWeeklyGoal(float weeklyGoal) {
		this.weeklyGoal = weeklyGoal;
	}

	/**
	 * @return the weeklySold
	 */
	public float getWeeklySold() {
		return weeklySold;
	}

	/**
	 * @param weeklySold the weeklySold to set
	 */
	public void setWeeklySold(float weeklySold) {
		this.weeklySold = weeklySold;
	}

	/**
	 * @return the monthlySold
	 */
	public float getMonthlySold() {
		return monthlySold;
	}

	/**
	 * @param monthlySold the monthlySold to set
	 */
	public void setMonthlySold(float monthlySold) {
		this.monthlySold = monthlySold;
	}

	/**
	 * @return the idEnterprise
	 */
	public long getIdEnterprise() {
		return idEnterprise;
	}

	/**
	 * @param idEnterprise the idEnterprise to set
	 */
	public void setIdEnterprise(long idEnterprise) {
		this.idEnterprise = idEnterprise;
	}

	/**
	 * @return the enterprise
	 */
	public Enterprise getEnterprise() {
		return enterprise;
	}

	/**
	 * @param enterprise the enterprise to set
	 */
	public void setEnterprise(Enterprise enterprise) {
		this.enterprise = enterprise;
	}
    
    
    
    
}
