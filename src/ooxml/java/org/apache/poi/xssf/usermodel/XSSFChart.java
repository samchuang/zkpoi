/* XSSFChart.java

	Purpose:
		
	Description:
		
	History:
		Oct 18, 2010 12:08:12 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.poi.xssf.usermodel;

import java.io.IOException;

import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTArea3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBar3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBubbleChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartSpace;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegendPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLine3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTOfPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPie3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTRadarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScatterChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStockChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSurface3DChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSurfaceChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTView3D;
import org.openxmlformats.schemas.drawingml.x2006.chart.ChartSpaceDocument;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.WsDrDocument;
import org.zkoss.poi.POIXMLDocumentPart;
import org.zkoss.poi.hssf.usermodel.HSSFChart.HSSFSeries;
import org.zkoss.poi.openxml4j.opc.PackagePart;
import org.zkoss.poi.openxml4j.opc.PackageRelationship;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.ChartInfo;
import org.zkoss.poi.ss.usermodel.ClientAnchor;

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
    
    public int getLegendPos() {
    	final CTLegend legend = _ctChart.getLegend();
    	final CTLegendPos pos = legend.getLegendPos();
    	return pos.getVal().intValue();
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

    public boolean isAutoTitleDeleted() {
    	final CTBoolean b = _ctChart.getAutoTitleDeleted();
    	return b != null ? b.getVal() : false; 
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
			final TitleInfo ti = getTitleInfo(ser);
			final CategoryInfo ci = getCategoryInfo(ser);
			final ValueInfo vi = getValueInfo(ser);
			final String tiref = ti == null ? null : ti.ref;
			final String tilit = ti == null ? null : ti.lit;
			final String ciref = ci == null ? null : ci.ref;
			final String[] cilits = ci == null ? null : ci.lits;
			final String viref = vi == null ? null : vi.ref;
			final String[] vilits = vi == null ? null : vi.lits;
			sers[idx] = new XSSFSeries(ser, tiref, tilit, ciref, cilits, viref, vilits);
		}
		return sers;
    }
	private TitleInfo getTitleInfo(CTBarSer ser) {
		if (ser != null) {
			final CTSerTx tx = ser.getTx();
			if (tx != null) {
				final CTStrRef strRef = tx.getStrRef();
				return new TitleInfo((strRef != null ? strRef.getF() : null), tx.getV());
			}
		}
		return null;
	}
	private CategoryInfo getCategoryInfo(CTBarSer ser) {
		if (ser != null) {
			final CTAxDataSource cat = ser.getCat();
			if (cat != null) {
				final CTStrRef strRef = cat.getStrRef();
				if (strRef != null) {
					final CTStrData strData = strRef.getStrCache();
					final String[] lits = getStrLits(strData);
					return new CategoryInfo(strRef.getF(), lits);
				}
				
				final CTNumRef numRef = cat.getNumRef();
				if (numRef != null) {
					final CTNumData numData = numRef.getNumCache();
					final String[] lits = getNumLits(numData);
					return new CategoryInfo(numRef.getF(), lits);
				}
				
				final CTStrData strData = cat.getStrLit();
				if (strData != null) {
					final String[] lits = getStrLits(strData);
					return new CategoryInfo(null, lits);
				}
				
				final CTNumData numData = cat.getNumLit();
				if (numData != null) {
					final String[] lits = getNumLits(numData);
					return new CategoryInfo(null, lits);
				}
			}
		}
		return null;
	}
	private ValueInfo getValueInfo(CTBarSer ser) {
		if (ser != null) {
			final CTNumDataSource numDS = ser.getVal();
			if (numDS != null) {
				final CTNumRef numRef = numDS.getNumRef();
				if (numRef != null) {
					final CTNumData numData = numRef.getNumCache();
					final String[] lits = getNumLits(numData);
					return new ValueInfo(numRef.getF(), lits);
				}
				final CTNumData numData = numDS.getNumLit();
				if (numData != null) {
					final String[] lits = getNumLits(numData);
					return new ValueInfo(null, lits);
				}
			}
		}
		return null;
	}
    
    //Pie
    private XSSFSeries[] preparePieSer(CTPieSer[] bsers) {
		final XSSFSeries[] sers = new XSSFSeries[bsers.length];
		for (int j = 0; j < bsers.length; ++j) {
			final CTPieSer ser = bsers[j];
			final int idx = (int) ser.getIdx().getVal();
			final TitleInfo ti = getTitleInfo(ser);
			final CategoryInfo ci = getCategoryInfo(ser);
			final ValueInfo vi = getValueInfo(ser);
			final String tiref = ti == null ? null : ti.ref;
			final String tilit = ti == null ? null : ti.lit;
			final String ciref = ci == null ? null : ci.ref;
			final String[] cilits = ci == null ? null : ci.lits;
			final String viref = vi == null ? null : vi.ref;
			final String[] vilits = vi == null ? null : vi.lits;
			sers[idx] = new XSSFSeries(ser, tiref, tilit, ciref, cilits, viref, vilits);
		}
		return sers;
    }
	private TitleInfo getTitleInfo(CTPieSer ser) {
		if (ser != null) {
			final CTSerTx tx = ser.getTx();
			if (tx != null) {
				final CTStrRef strRef = tx.getStrRef();
				return new TitleInfo((strRef != null ? strRef.getF() : null), tx.getV());
			}
		}
		return null;
	}
	private CategoryInfo getCategoryInfo(CTPieSer ser) {
		if (ser != null) {
			final CTAxDataSource cat = ser.getCat();
			if (cat != null) {
				final CTStrRef strRef = cat.getStrRef();
				if (strRef != null) {
					final CTStrData strData = strRef.getStrCache();
					final String[] lits = getStrLits(strData);
					return new CategoryInfo(strRef.getF(), lits);
				}
				
				final CTNumRef numRef = cat.getNumRef();
				if (numRef != null) {
					final CTNumData numData = numRef.getNumCache();
					final String[] lits = getNumLits(numData);
					return new CategoryInfo(numRef.getF(), lits);
				}
				
				final CTStrData strData = cat.getStrLit();
				if (strData != null) {
					final String[] lits = getStrLits(strData);
					return new CategoryInfo(null, lits);
				}
				
				final CTNumData numData = cat.getNumLit();
				if (numData != null) {
					final String[] lits = getNumLits(numData);
					return new CategoryInfo(null, lits);
				}
			}
		}
		return null;
	}
	private ValueInfo getValueInfo(CTPieSer ser) {
		if (ser != null) {
			final CTNumDataSource numDS = ser.getVal();
			if (numDS != null) {
				final CTNumRef numRef = numDS.getNumRef();
				if (numRef != null) {
					final CTNumData numData = numRef.getNumCache();
					final String[] lits = getNumLits(numData);
					return new ValueInfo(numRef.getF(), lits);
				}
				final CTNumData numData = numDS.getNumLit();
				if (numData != null) {
					final String[] lits = getNumLits(numData);
					return new ValueInfo(null, lits);
				}
			}
		}
		return null;
	}

    //Line
    private XSSFSeries[] prepareLineSer(CTLineSer[] lsers) {
		final XSSFSeries[] sers = new XSSFSeries[lsers.length];
		for (int j = 0; j < lsers.length; ++j) {
			final CTLineSer ser = lsers[j];
			final int idx = (int) ser.getIdx().getVal();
			final TitleInfo ti = getTitleInfo(ser);
			final CategoryInfo ci = getCategoryInfo(ser);
			final ValueInfo vi = getValueInfo(ser);
			final String tiref = ti == null ? null : ti.ref;
			final String tilit = ti == null ? null : ti.lit;
			final String ciref = ci == null ? null : ci.ref;
			final String[] cilits = ci == null ? null : ci.lits;
			final String viref = vi == null ? null : vi.ref;
			final String[] vilits = vi == null ? null : vi.lits;
			sers[idx] = new XSSFSeries(ser, tiref, tilit, ciref, cilits, viref, vilits);
		}
		return sers;
    }
	private TitleInfo getTitleInfo(CTLineSer ser) {
		if (ser != null) {
			final CTSerTx tx = ser.getTx();
			if (tx != null) {
				final CTStrRef strRef = tx.getStrRef();
				return new TitleInfo((strRef != null ? strRef.getF() : null), tx.getV());
			}
		}
		return null;
	}
	private CategoryInfo getCategoryInfo(CTLineSer ser) {
		if (ser != null) {
			final CTAxDataSource cat = ser.getCat();
			if (cat != null) {
				final CTStrRef strRef = cat.getStrRef();
				if (strRef != null) {
					final CTStrData strData = strRef.getStrCache();
					final String[] lits = getStrLits(strData);
					return new CategoryInfo(strRef.getF(), lits);
				}
				
				final CTNumRef numRef = cat.getNumRef();
				if (numRef != null) {
					final CTNumData numData = numRef.getNumCache();
					final String[] lits = getNumLits(numData);
					return new CategoryInfo(numRef.getF(), lits);
				}
				
				final CTStrData strData = cat.getStrLit();
				if (strData != null) {
					final String[] lits = getStrLits(strData);
					return new CategoryInfo(null, lits);
				}
				
				final CTNumData numData = cat.getNumLit();
				if (numData != null) {
					final String[] lits = getNumLits(numData);
					return new CategoryInfo(null, lits);
				}
			}
		}
		return null;
	}
	private ValueInfo getValueInfo(CTLineSer ser) {
		if (ser != null) {
			final CTNumDataSource numDS = ser.getVal();
			if (numDS != null) {
				final CTNumRef numRef = numDS.getNumRef();
				if (numRef != null) {
					final CTNumData numData = numRef.getNumCache();
					final String[] lits = getNumLits(numData);
					return new ValueInfo(numRef.getF(), lits);
				}
				final CTNumData numData = numDS.getNumLit();
				if (numData != null) {
					final String[] lits = getNumLits(numData);
					return new ValueInfo(null, lits);
				}
			}
		}
		return null;
	}
    
	private String[] getStrLits(CTStrData strData) {
		String[] lits = null;
		if (strData != null) {
			final CTStrVal[] strVals = strData.getPtArray();
			lits = new String[strVals.length];
			for(int j = 0; j < strVals.length; ++j) {
				lits[j] = strVals[j].getV();
			}
		}
		return lits;
	}
	private String[] getNumLits(CTNumData numData) {
		String[] lits = null;
		if (numData != null) {
			final CTNumVal[] numVals = numData.getPtArray();
			lits = new String[numVals.length];
			for(int j= 0; j < numVals.length; ++j) {
				lits[j] = numVals[j].getV();
			}
		}
		return lits;
	}
    private static class TitleInfo {
    	private final String ref;
    	private final String lit;
    	private TitleInfo(String ref, String lit) {
    		this.ref = ref;
    		this.lit = lit;
    	}
    }
	private static class CategoryInfo {
		private final String ref;
		private final String[] lits;
		private CategoryInfo(String ref, String[] lits) {
			this.ref = ref;
			this.lits = lits;
		}
	}
	private static class ValueInfo {
		private final String ref;
		private final String[] lits;
		private ValueInfo(String ref, String[] lits) {
			this.ref = ref;
			this.lits = lits;
		}
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
		private final String[] _labels; //literal labels(categories)
		private final String _title; //literal title
		private final Number[] _values; //literal values
		
		public XSSFSeries(XmlObject ser, String titleRef, String title, String catRef, String[] labels, String valRef, String[] values) {
			_ser = ser;
			_titleRef = titleRef;
			if (catRef != null) {
				if (catRef.startsWith("(") && catRef.endsWith(")")) {
					catRef = catRef.substring(1, catRef.length() - 1);
				}
			}
			_catRef = catRef;
			if (valRef != null) {
				if (valRef.startsWith("(") && valRef.endsWith(")")) {
					valRef = valRef.substring(1, valRef.length() - 1);
				}
			}
			_valRef = valRef;
			_title = title;
			_labels = labels;
			Double[] vals = null;
			if (values != null) {
				vals = new Double[values.length];
				for(int j = 0; j < values.length; ++j) {
					Double db = new Double(0.0);
					try {
						db = new Double(values[j]);
					} catch(NumberFormatException ex) {
						//ignore
					}
					vals[j++] = db; 
				}
			}
			_values = vals;
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
		
		public String getTitle() {
			return _title;
		}
		
		public String[] getLabels() {
			return _labels;
		}
		
		public Number[] getValues() {
			return _values;
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
