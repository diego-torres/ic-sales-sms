package entities;

public class Director
{
    private long id;
    private boolean sms = false;
    private String name;
    private String email;
    private String cellPhone;
    private long idEnterprise;

    private Enterprise enterprise;

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
