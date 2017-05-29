package com.googlesheets.poc.model;

import java.util.List;

public class Response {
	
	private String status;
	private String customerName;
	private String redirectUril;
	private List<GSheet> sheets;

	/**
	 * @return the status
	 */
	public String getStatus() {
		return status;
	}

	/**
	 * @param status the status to set
	 */
	public void setStatus(String status) {
		this.status = status;
	}


	/**
	 * @return the customerName
	 */
	public String getCustomerName() {
		return customerName;
	}

	/**
	 * @return the sheets
	 */
	public List<GSheet> getSheets() {
		return sheets;
	}

	/**
	 * @param sheets the sheets to set
	 */
	public void setSheets(List<GSheet> sheets) {
		this.sheets = sheets;
	}

	/**
	 * @param customerName the customerName to set
	 */
	public void setCustomerName(String customerName) {
		this.customerName = customerName;
	}

	/**
	 * @return the redirectUril
	 */
	public String getRedirectUril() {
		return redirectUril;
	}

	/**
	 * @param redirectUril the redirectUril to set
	 */
	public void setRedirectUril(String redirectUril) {
		this.redirectUril = redirectUril;
	}
}
