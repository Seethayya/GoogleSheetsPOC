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
public class GRow {

	private List<String> columns;

	/**
	 * @return the columns
	 */
	public List<String> getColumns() {
		if (columns == null)
			columns = new ArrayList<String>();
		return columns;
	}

	/**
	 * @param columns the columns to set
	 */
	public void setColumns(List<String> columns) {
		this.columns = columns;
	}
	
	
}
