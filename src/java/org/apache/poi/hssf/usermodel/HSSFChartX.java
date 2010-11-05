/* HSSFChartX.java

	Purpose:
		
	Description:
		
	History:
		Oct 14, 2010 7:04:50 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.poi.hssf.usermodel;

import org.zkoss.poi.hssf.usermodel.HSSFChart;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.ChartInfo;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.util.CellRangeAddressBase;

/**
 * Represents a escher chart.
 * 
 * @author henrichen
 */
public class HSSFChartX extends HSSFSimpleShape implements Chart {
	private HSSFChart _chart;
	private String _name;
    HSSFPatriarch _patriarch;  // TODO make private
	
	public HSSFChartX(HSSFShape parent, HSSFAnchor anchor) {
        super( parent, anchor );
        setShapeType(OBJECT_TYPE_CHART);
	}

	public ChartInfo getChartInfo() {
		return _chart;
	}

	public void setChart(HSSFChart chart) {
		this._chart = chart;
	}

	public String getName() {
		return _name;
	}

	public void setName(String name) {
		this._name = name;
	}

	@Override
	public ClientAnchor getPreferredSize() {
        return (ClientAnchor)getAnchor();
	}
}
