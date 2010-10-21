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
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ChartInfo;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChartSpace;
import org.openxmlformats.schemas.drawingml.x2006.chart.ChartSpaceDocument;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.WsDrDocument;

/**
 * XSSFChart information.
 * @author henrichen
 *
 */
public class XSSFChart extends POIXMLDocumentPart implements ChartInfo {
	private CTChartSpace _ctChartSpace;

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
    }
}
