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

import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.charts.AbstractCategoryDataSerie;
import org.zkoss.poi.ss.usermodel.charts.CategoryData;
import org.zkoss.poi.ss.usermodel.charts.ChartAxis;
import org.zkoss.poi.ss.usermodel.charts.ChartDataSource;
import org.zkoss.poi.ss.usermodel.charts.ChartTextSource;
import org.zkoss.poi.ss.usermodel.charts.CategoryDataSerie;
import org.zkoss.poi.util.Beta;
import org.zkoss.poi.xssf.usermodel.XSSFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.util.ArrayList;
import java.util.List;

/**
 * Represents DrawingML bar chart.
 *
 * @author henrichen@zkoss.org
 */
@Beta
public class XSSFBarChartData implements CategoryData {

	private CTBarChart ctBarChart;
    /**
     * List of all data series.
     */
    private List<CategoryDataSerie> series;

    public XSSFBarChartData() {
        series = new ArrayList<CategoryDataSerie>();
    }

    public XSSFBarChartData(XSSFChart chart) {
    	this();
    	final CTPlotArea plotArea = chart.getCTChart().getPlotArea();
    	
    	//Bar
    	final CTBarChart[] plotCharts = plotArea.getBarChartArray();
    	if (plotCharts != null && plotCharts.length > 0) {
    		ctBarChart = plotCharts[0];
    	}
    }
    
    /**
     * Package private PieChartSerie implementation.
     */
    static class Serie extends AbstractCategoryDataSerie {
        protected Serie(int id, int order, ChartTextSource title,
				ChartDataSource<?> cats, ChartDataSource<? extends Number> vals) {
			super(id, order, title, cats, vals);
		}

		protected void addToChart(CTBarChart ctBarChart) {
            CTBarSer barSer = ctBarChart.addNewSer();
            barSer.addNewIdx().setVal(this.id);
            barSer.addNewOrder().setVal(this.order);

            CTSerTx tx = barSer.addNewTx();
            XSSFChartUtil.buildSerTx(tx, title);
            
            CTAxDataSource cats = barSer.addNewCat();
            XSSFChartUtil.buildAxDataSource(cats, categories);

            CTNumDataSource vals = barSer.addNewVal();
            XSSFChartUtil.buildNumDataSource(vals, values);
        }
    }

    public CategoryDataSerie addSerie(ChartTextSource title, ChartDataSource<?> cats,
                                      ChartDataSource<? extends Number> vals) {
        if (!vals.isNumeric()) {
            throw new IllegalArgumentException("Bar data source must be numeric.");
        }
        int numOfSeries = series.size();
        Serie newSerie = new Serie(numOfSeries, numOfSeries, title, cats, vals);
        series.add(newSerie);
        return newSerie;
    }

    public void fillChart(Chart chart, ChartAxis... axis) {
        if (!(chart instanceof XSSFChart)) {
            throw new IllegalArgumentException("Chart must be instance of XSSFChart");
        }

        if (ctBarChart == null) {
	        XSSFChart xssfChart = (XSSFChart) chart;
	        CTPlotArea plotArea = xssfChart.getCTChart().getPlotArea();
	        ctBarChart = plotArea.addNewBarChart();
        
	        //TODO setup other properties of barChart
	
	        for (CategoryDataSerie s : series) {
	            ((Serie)s).addToChart(ctBarChart);
	        }
        }
    }

    public List<? extends CategoryDataSerie> getSeries() {
        return series;
    }
}
