package gateway;

public class GatewayResponse {
	private Object result = null;
	private boolean errorThrown = false;
	private String responseDescription = "";

	/**
	 * @return the result
	 */
	public Object getResult() {
		return result;
	}

	/**
	 * @param result
	 *            the result to set
	 */
	public void setResult(Object result) {
		this.result = result;
	}

	/**
	 * @return the errorThrown
	 */
	public boolean isErrorThrown() {
		return errorThrown;
	}

	/**
	 * @param errorThrown
	 *            the errorThrown to set
	 */
	public void setErrorThrown(boolean errorThrown) {
		this.errorThrown = errorThrown;
	}

	/**
	 * @return the responseDescription
	 */
	public String getResponseDescription() {
		return responseDescription;
	}

	/**
	 * @param responseDescription
	 *            the responseDescription to set
	 */
	public void setResponseDescription(String responseDescription) {
		this.responseDescription = responseDescription;
	}

}
