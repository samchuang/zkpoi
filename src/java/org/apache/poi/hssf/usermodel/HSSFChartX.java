/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */
package org.zkoss.poi.hssf.usermodel;

import org.zkoss.poi.hssf.usermodel.HSSFChart;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.ZssChartX;
import org.zkoss.poi.ss.usermodel.ChartInfo;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.util.CellRangeAddressBase;

/**
 * Represents a escher chart.
 * 
 * @author henrichen@zkoss.org
 */
public class HSSFChartX extends HSSFSimpleShape implements ZssChartX {
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

	@Override
	public Chart getChart() {
		return null; //return _chart;
	}

	@Override
	public String getChartId() {
		return getName();
	}

	//20111110, henrichen@zkoss.org: update chart anchor
	@Override
	public void setClientAnchor(ClientAnchor newanchor) {
        HSSFClientAnchor anchor = (HSSFClientAnchor)getAnchor();
    	anchor.setCol1(newanchor.getCol1());
    	anchor.setCol2(newanchor.getCol2());
    	anchor.setDx1(newanchor.getDx1());
    	anchor.setDx2(newanchor.getDx2());
    	anchor.setDy1(newanchor.getDy1());
    	anchor.setDy2(newanchor.getDy2());
    	anchor.setRow1(newanchor.getRow1());
    	anchor.setRow2(newanchor.getRow2());
	}
}
