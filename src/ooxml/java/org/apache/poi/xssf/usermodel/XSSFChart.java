/* XSSFChart.java

	Purpose:
		
	Description:
		
	History:
		Oct 18, 2010 12:08:12 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.apache.poi.xssf.usermodel;

import java.io.IOException;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.HSSFChart.HSSFSeries;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ChartInfo;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTArea3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBar3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBubbleChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartSpace;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLine3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTOfPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPie3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTRadarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScatterChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStockChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSurface3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSurfaceChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTView3D;
import org.openxmlformats.schemas.drawingml.x2006.chart.ChartSpaceDocument;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.WsDrDocument;

/**
 * XSSFChart information.
 * @author henrichen
 *
 */
public class XSSFChart extends POIXMLDocumentPart implements ChartInfo {
	private CTChartSpace _ctChartSpace;
	private CTChart _ctChart;

    /**
     * Construct a SpreadsheetML drawing from a package part
     *
     * @param part the package part holding the drawing data,
     * the content type must be <code>application/vnd.openxmlformats-officedocument.drawing+xml</code>
     * @param rel  the package relationship holding this drawing,
     * the relationship type must be http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing
     */
    protected XSSFChart(PackagePart part, PackageRelationship rel) throws IOException, XmlException {
        super(part, rel);
        _ctChartSpace = ChartSpaceDocument.Factory.parse(part.getInputStream()).getChartSpace();
        _ctChart = _ctChartSpace.getChart();
    }
    
    public String getChartTitle() {
    	final CTTitle title = _ctChart.getTitle();
    	if (title != null) {
    		final CTTx tx = title.getTx();
	    	if (tx != null) {
		    	final CTTextParagraph[] paras = tx.getRich().getPArray();
		    	final CTRegularTextRun[] runs = paras[0].getRArray();
		    	return runs[0].getT();
	    	}
    	}
    	return null;
    }
    
    public boolean getChart3D() {
    	final CTView3D view3D = _ctChart.getView3D();
    	return view3D != null;
    }
    
    public XSSFChartType getType() {
    	final XmlObject chartObj = getChartObj();
    	if (chartObj instanceof CTArea3DChart || chartObj instanceof CTAreaChart) {
    		return XSSFChartType.Area;
    	}
    	
    	if (chartObj instanceof CTBar3DChart || chartObj instanceof CTBarChart) {
    		return XSSFChartType.Bar;
    	}

    	if (chartObj instanceof CTBubbleChart) {
    		return XSSFChartType.Bubble;
    	}

    	if (chartObj instanceof CTDoughnutChart) {
    		return XSSFChartType.Donut;
    	}

    	if (chartObj instanceof CTLine3DChart || chartObj instanceof CTLineChart) {
    		return XSSFChartType.Line;
    	}

    	if (chartObj instanceof CTOfPieChart || chartObj instanceof CTPie3DChart || chartObj instanceof CTPieChart) {
    		return XSSFChartType.Pie;
    	}

    	if (chartObj instanceof CTRadarChart) {
    		return XSSFChartType.Radar;
    	}

    	if (chartObj instanceof CTScatterChart) {
    		return XSSFChartType.Scatter;
    	}
    	
    	if (chartObj instanceof CTStockChart) {
    		return XSSFChartType.Stock;
    	}

    	if (chartObj instanceof CTSurfaceChart || chartObj instanceof CTSurface3DChart) {
    		return XSSFChartType.Surface;
    	}
    	
    	return XSSFChartType.Unknown;
    }
    
    private XmlObject getChartObj() {
    	final CTPlotArea plotArea = _ctChart.getPlotArea();
    	//Area3D
    	final CTArea3DChart[] area3ds = plotArea.getArea3DChartArray();
    	if (area3ds != null && area3ds.length > 0) {
    		return area3ds[0];
    	}
    	
    	//Area
    	final CTAreaChart[] areas = plotArea.getAreaChartArray();
    	if (areas != null && areas.length > 0) {
    		return areas[0];
    	}

    	//Bar3D
    	final CTBar3DChart[] bar3ds = plotArea.getBar3DChartArray();
    	if (bar3ds != null && bar3ds.length > 0) {
    		return bar3ds[0];
    	}
    	
    	//Bar
    	final CTBarChart[] bars = plotArea.getBarChartArray();
    	if (bars != null && bars.length > 0) {
    		return bars[0];
    	}
    	
    	//Bubble
    	final CTBubbleChart[] bubbles = plotArea.getBubbleChartArray();
    	if (bubbles != null && bubbles.length > 0) {
    		return bubbles[0];
    	}

    	//Doughnut
    	final CTDoughnutChart[] donuts = plotArea.getDoughnutChartArray();
    	if (donuts != null && donuts.length > 0) {
    		return donuts[0];
    	}
    	
    	//Line3D
    	final CTLine3DChart[] line3ds = plotArea.getLine3DChartArray();
    	if (line3ds != null && line3ds.length > 0) {
    		return line3ds[0];
    	}
    	
    	//Line
    	final CTLineChart[] lines = plotArea.getLineChartArray();
    	if (lines != null && lines.length > 0) {
    		return lines[0];
    	}
    	
    	//OfPie
    	final CTOfPieChart[] ofpies = plotArea.getOfPieChartArray();
    	if (ofpies != null && ofpies.length > 0) {
    		return ofpies[0];
    	}
    	
    	//Pie3D
    	final CTPie3DChart[] pie3ds = plotArea.getPie3DChartArray();
    	if (pie3ds != null && pie3ds.length > 0) {
    		return pie3ds[0];
    	}
    	
    	//Pie
    	final CTPieChart[] pies = plotArea.getPieChartArray();
    	if (pies != null && pies.length > 0) {
    		return pies[0];
    	}

    	//Radar
    	final CTRadarChart[] radars = plotArea.getRadarChartArray();
    	if (radars != null && radars.length > 0) {
    		return radars[0];
    	}

    	//Scatter
    	final CTScatterChart[] scatters = plotArea.getScatterChartArray();
    	if (scatters != null && scatters.length > 0) {
    		return scatters[0];
    	}
    	
    	//Stock
    	final CTStockChart[] stocks = plotArea.getStockChartArray();
    	if (stocks != null && stocks.length > 0) {
    		return stocks[0];
    	}

    	//Surface3D
    	final CTSurface3DChart[] surface3ds = plotArea.getSurface3DChartArray();
    	if (surface3ds != null && surface3ds.length > 0) {
    		return surface3ds[0];
    	}
    	
    	//Surface
    	final CTSurfaceChart[] surfaces = plotArea.getSurfaceChartArray();
    	if (surfaces != null && surfaces.length > 0) {
    		return surfaces[0];
    	}
    	
    	return null;
    }
    
    public XSSFSeries[] getSeries() {
    	//TODO, more Chart type
    	final XmlObject chartObj = getChartObj();
   		if (chartObj instanceof CTBar3DChart) {
   			final CTBar3DChart chart = (CTBar3DChart) chartObj;
   			final CTBarSer[] bsers = chart.getSerArray();
   			return prepareBarSer(bsers);
    	} else if (chartObj instanceof CTBarChart) {
   			final CTBarChart chart = (CTBarChart) chartObj;
   			final CTBarSer[] bsers = chart.getSerArray();
   			return prepareBarSer(bsers);
    	} else if (chartObj instanceof CTPieChart) {
    		final CTPieChart chart = (CTPieChart) chartObj;
    		final CTPieSer[] psers = chart.getSerArray();
    		return preparePieSer(psers);
    	} else if (chartObj instanceof CTPie3DChart) {
    		final CTPie3DChart chart = (CTPie3DChart) chartObj;
    		final CTPieSer[] psers = chart.getSerArray();
    		return preparePieSer(psers);
    	} else if (chartObj instanceof CTDoughnutChart) {
    		final CTDoughnutChart chart = (CTDoughnutChart) chartObj;
    		final CTPieSer[] psers = chart.getSerArray();
    		return preparePieSer(psers);
    	} else if (chartObj instanceof CTLineChart) {
    		final CTLineChart chart = (CTLineChart) chartObj;
    		final CTLineSer[] lsers = chart.getSerArray();
    		return prepareLineSer(lsers);
    	} else if (chartObj instanceof CTLine3DChart) {
    		final CTLine3DChart chart = (CTLine3DChart) chartObj;
    		final CTLineSer[] lsers = chart.getSerArray();
    		return prepareLineSer(lsers);
/*    	} else if (chartObj instanceof CTAreaChart) {
    		final CTAreaChart chart = (CTAreaChart) chartObj;
    		final CTAreaSer[] asers = chart.getSerArray();
    		return prepareAreaSer(asers);
    	} else if (chartObj instanceof CTArea3DChart) {
    		final CTArea3DChart chart = (CTArea3DChart) chartObj;
    		final CTAreaSer[] asers = chart.getSerArray();
    		return prepareAreaSer(asers);
 */   	}
   		return null;
    }
    
    //Bar
    private XSSFSeries[] prepareBarSer(CTBarSer[] bsers) {
   			final XSSFSeries[] sers = new XSSFSeries[bsers.length];
   			for (int j = 0; j < bsers.length; ++j) {
   				final CTBarSer ser = bsers[j];
   				final int idx = (int) ser.getIdx().getVal();
   				sers[idx] = new XSSFSeries(ser, getTitleRef(ser), getCategoryRef(ser), getValueRef(ser));
   			}
   			return sers;
    }
	private String getTitleRef(CTBarSer ser) {
		return ser.getTx().getStrRef().getF();
	}
	private String getCategoryRef(CTBarSer ser) {
		final CTAxDataSource cat = ser.getCat();
		final CTStrRef strRef = cat.getStrRef(); 
		return strRef != null ? strRef.getF() : cat.getNumRef().getF(); //could be string or number 
	}
	private String getValueRef(CTBarSer ser) {
		return ser.getVal().getNumRef().getF();
	}
    
    //Pie
    private XSSFSeries[] preparePieSer(CTPieSer[] bsers) {
			final XSSFSeries[] sers = new XSSFSeries[bsers.length];
			for (int j = 0; j < bsers.length; ++j) {
				final CTPieSer ser = bsers[j];
				final int idx = (int) ser.getIdx().getVal();
				sers[idx] = new XSSFSeries(ser, getTitleRef(ser), getCategoryRef(ser), getValueRef(ser));
			}
			return sers;
    }
	private String getTitleRef(CTPieSer ser) {
		return ser.getTx().getStrRef().getF();
	}
	private String getCategoryRef(CTPieSer ser) {
		final CTAxDataSource cat = ser.getCat();
		final CTStrRef strRef = cat.getStrRef(); 
		return strRef != null ? strRef.getF() : cat.getNumRef().getF(); //could be string or number 
	}
	private String getValueRef(CTPieSer ser) {
		return ser.getVal().getNumRef().getF();
	}

    //Line
    private XSSFSeries[] prepareLineSer(CTLineSer[] lsers) {
   			final XSSFSeries[] sers = new XSSFSeries[lsers.length];
   			for (int j = 0; j < lsers.length; ++j) {
   				final CTLineSer ser = lsers[j];
   				final int idx = (int) ser.getIdx().getVal();
   				sers[idx] = new XSSFSeries(ser, getTitleRef(ser), getCategoryRef(ser), getValueRef(ser));
   			}
   			return sers;
    }
	private String getTitleRef(CTLineSer ser) {
		return ser.getTx().getStrRef().getF();
	}
	private String getCategoryRef(CTLineSer ser) {
		final CTAxDataSource cat = ser.getCat();
		final CTStrRef strRef = cat.getStrRef(); 
		return strRef != null ? strRef.getF() : cat.getNumRef().getF(); //could be string or number 
	}
	private String getValueRef(CTLineSer ser) {
		return ser.getVal().getNumRef().getF();
	}
    
	/**
	 * A series in a chart.
	 */
	public class XSSFSeries {
    	//TODO, more Chart type
		private final XmlObject _ser; //might be CTBarSer, CTPieSer
		private final String _titleRef; //title reference formula(e.g. B1:B1)
		private final String _catRef; //category reference formula(e.g. A2:A4)
		private final String _valRef; //value reference formula(e.g. B2:B4)
		
		public XSSFSeries(XmlObject ser, String titleRef, String catRef, String valRef) {
			_ser = ser;
			_titleRef = titleRef;
			_catRef = catRef;
			_valRef = valRef;
		}
		public String getTitleRef() {
			return _titleRef;
		}
 
		public String getCategoryRef() {
			return _catRef;
		}
		
		public String getValueRef() {
			return _valRef;
		}
	}

	public enum XSSFChartType {
		Area,
		Bar,
		Bubble,
		Donut,
		Line,
		Pie,
		Radar,
		Scatter,
		Stock,
		Surface,
		Unknown
	}
}
