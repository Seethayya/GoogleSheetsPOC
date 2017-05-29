/**
 * 
 */
package com.googlesheets.poc.model;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Seethayya
 *
 */
public class GSheet {
	
	private String name;

	private List<GRow> rows;

	
	
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
	 * @return the rows
	 */
	public List<GRow> getRows() {
		if (rows == null)
			rows = new ArrayList<GRow>();
		return rows;
	}

	/**
	 * @param rows the rows to set
	 */
	public void setRows(List<GRow> rows) {
		this.rows = rows;
	}
		
		
}
