/* FilterColumn.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		May 7, 2011 6:28:58 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/


package org.zkoss.poi.ss.usermodel;

import java.util.List;

/**
 * Represent a filtered column.
 * @author henrichen
 *
 */
public interface FilterColumn {
	/**
	 * Returns the column number of this FilterColumn. 
	 * @return the column number of this FilterColumn.
	 */
	public int getColumn();
	
	/**
	 * Returns the filter String list of this FilterColumn.
	 * @return the filter String list of this FilterColumn.
	 */
	public List<String> getFilters();
}
