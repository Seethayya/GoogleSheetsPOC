/**
 * 
 */
package com.googlesheets.poc.model;

/**
 * @author Seethayya
 *
 */
public class GApplicationCredential {

	private String clientId;
	private String projctId;
	private String authUri;
	private String toeknUri;
	private String athProviderUrl;
	private String secretId;
	private String redirectUri;
	/**
	 * @return the clientId
	 */
	public String getClientId() {
		return clientId;
	}
	/**
	 * @param clientId
	 * @param projctId
	 * @param authUri
	 * @param toeknUri
	 * @param athProviderUrl
	 * @param secretId
	 */
	public GApplicationCredential(String clientId, String projctId, String authUri, String toeknUri,
			String athProviderUrl, String secretId, String redirectUri) {
		super();
		this.clientId = clientId;
		this.projctId = projctId;
		this.authUri = authUri;
		this.toeknUri = toeknUri;
		this.athProviderUrl = athProviderUrl;
		this.secretId = secretId;
		this.redirectUri=redirectUri;
	}
	/**
	 * @param clientId the clientId to set
	 */
	public void setClientId(String clientId) {
		this.clientId = clientId;
	}
	/**
	 * @return the projctId
	 */
	public String getProjctId() {
		return projctId;
	}
	/**
	 * @param projctId the projctId to set
	 */
	public void setProjctId(String projctId) {
		this.projctId = projctId;
	}
	/**
	 * @return the authUri
	 */
	public String getAuthUri() {
		return authUri;
	}
	/**
	 * @param authUri the authUri to set
	 */
	public void setAuthUri(String authUri) {
		this.authUri = authUri;
	}
	/**
	 * @return the toeknUri
	 */
	public String getToeknUri() {
		return toeknUri;
	}
	/**
	 * @param toeknUri the toeknUri to set
	 */
	public void setToeknUri(String toeknUri) {
		this.toeknUri = toeknUri;
	}
	/**
	 * @return the athProviderUrl
	 */
	public String getAthProviderUrl() {
		return athProviderUrl;
	}
	/**
	 * @param athProviderUrl the athProviderUrl to set
	 */
	public void setAthProviderUrl(String athProviderUrl) {
		this.athProviderUrl = athProviderUrl;
	}
	/**
	 * @return the secretId
	 */
	public String getSecretId() {
		return secretId;
	}
	/**
	 * @param secretId the secretId to set
	 */
	public void setSecretId(String secretId) {
		this.secretId = secretId;
	}
	/**
	 * @return the redirectUri
	 */
	public String getRedirectUri() {
		return redirectUri;
	}
	/**
	 * @param redirectUri the redirectUri to set
	 */
	public void setRedirectUri(String redirectUri) {
		this.redirectUri = redirectUri;
	}
}
