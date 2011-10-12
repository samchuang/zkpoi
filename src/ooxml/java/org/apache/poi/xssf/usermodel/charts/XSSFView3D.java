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
package org.zkoss.poi.xssf.usermodel.charts;

import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTView3D;
import org.zkoss.poi.ss.usermodel.charts.View3D;
import org.zkoss.poi.util.Beta;
import org.zkoss.poi.xssf.usermodel.XSSFChart;

/**
 * Represents a SpreadsheetML chart 3D view information.
 * 
 * @author henrichen@zkoss.org
 */
@Beta
public class XSSFView3D implements View3D {
	/**
	 * Underlying CTView3D bean
	 */
	private CTView3D view3d;

	/**
	 * Create a new SpreadsheetML chart legend
	 */
	public XSSFView3D(XSSFChart chart) {
		CTChart ctChart = chart.getCTChart();
		this.view3d = (ctChart.isSetView3D()) ?
			ctChart.getView3D() :
			ctChart.addNewView3D();
	}
}
