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

package org.zkoss.poi.xssf.usermodel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import javax.xml.namespace.QName;

import org.zkoss.poi.POIXMLDocumentPart;
import org.zkoss.poi.POIXMLException;
import org.zkoss.poi.hssf.record.PasswordRecord;
import org.zkoss.poi.hssf.util.PaneInformation;
import org.zkoss.poi.openxml4j.exceptions.InvalidFormatException;
import org.zkoss.poi.openxml4j.exceptions.PartAlreadyExistsException;
import org.zkoss.poi.openxml4j.opc.PackagePart;
import org.zkoss.poi.openxml4j.opc.PackageRelationship;
import org.zkoss.poi.openxml4j.opc.PackageRelationshipCollection;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.FormulaShifter;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellRange;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.DataValidation;
import org.zkoss.poi.ss.usermodel.DataValidationHelper;
import org.zkoss.poi.ss.usermodel.FilterColumn;
import org.zkoss.poi.ss.usermodel.Footer;
import org.zkoss.poi.ss.usermodel.Header;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.ss.util.CellRangeAddressList;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.poi.ss.util.SSCellRange;
import org.zkoss.poi.ss.util.SheetUtil;
import org.zkoss.poi.util.HexDump;
import org.zkoss.poi.util.Internal;
import org.zkoss.poi.util.POILogFactory;
import org.zkoss.poi.util.POILogger;
import org.zkoss.poi.xssf.model.CommentsTable;
import org.zkoss.poi.xssf.usermodel.helpers.ColumnHelper;
import org.zkoss.poi.xssf.usermodel.helpers.XSSFRowShifter;
import org.zkoss.poi.xssf.usermodel.XSSFAutoFilter.XSSFFilterColumn;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.officeDocument.x2006.relationships.STRelationshipId;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBreak;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCellFormula;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCols;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCommentList;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilterColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTHeaderFooter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTLegacyDrawing;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMergeCell;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMergeCells;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTOutlinePr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageBreak;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageMargins;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetUpPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPane;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPrintOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSelection;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetCalcPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetFormatPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetProtection;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetView;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetViews;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTablePart;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableParts;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellFormulaType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPane;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPaneState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnsignedShortHex;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorksheetDocument;

/**
 * High level representation of a SpreadsheetML worksheet.
 *
 * <p>
 * Sheets are the central structures within a workbook, and are where a user does most of his spreadsheet work.
 * The most common type of sheet is the worksheet, which is represented as a grid of cells. Worksheet cells can
 * contain text, numbers, dates, and formulas. Cells can also be formatted.
 * </p>
 */
public class XSSFSheet extends POIXMLDocumentPart implements Sheet {
    private static final POILogger logger = POILogFactory.getLogger(XSSFSheet.class);

    //TODO make the two variable below private!
    protected CTSheet sheet;
    protected CTWorksheet worksheet;

    private TreeMap<Integer, XSSFRow> _rows;
    private List<XSSFHyperlink> hyperlinks;
    private ColumnHelper columnHelper;
    private CommentsTable sheetComments;
    
    
        
	/**
     * cache of master shared formulas in this sheet.
     * Master shared formula is the first formula in a group of shared formulas is saved in the f element.
     */
    private Map<Integer, CTCellFormula> sharedFormulas;
    private TreeMap<String,XSSFTable> tables;
    private List<CellRangeAddress> arrayFormulas;
    private XSSFDataValidationHelper dataValidationHelper;    

    /**
     * Creates new XSSFSheet   - called by XSSFWorkbook to create a sheet from scratch.
     *
     * @see org.zkoss.poi.xssf.usermodel.XSSFWorkbook#createSheet()
     */
    protected XSSFSheet() {
        super();
        dataValidationHelper = new XSSFDataValidationHelper(this);
        onDocumentCreate();
    }

    /**
     * Creates an XSSFSheet representing the given package part and relationship.
     * Should only be called by XSSFWorkbook when reading in an exisiting file.
     *
     * @param part - The package part that holds xml data represenring this sheet.
     * @param rel - the relationship of the given package part in the underlying OPC package
     */
    protected XSSFSheet(PackagePart part, PackageRelationship rel) {
        super(part, rel);
        dataValidationHelper = new XSSFDataValidationHelper(this);
    }

    /**
     * Returns the parent XSSFWorkbook
     *
     * @return the parent XSSFWorkbook
     */
    public XSSFWorkbook getWorkbook() {
        return (XSSFWorkbook)getParent();
    }

    /**
     * Initialize worksheet data when reading in an exisiting file.
     */
    @Override
    protected void onDocumentRead() {
        try {
            read(getPackagePart().getInputStream());
        } catch (IOException e){
            throw new POIXMLException(e);
        }
    }

    protected void read(InputStream is) throws IOException {
        try {
            worksheet = WorksheetDocument.Factory.parse(is).getWorksheet();
        } catch (XmlException e){
            throw new POIXMLException(e);
        }

        initRows(worksheet);
        columnHelper = new ColumnHelper(worksheet);

        // Look for bits we're interested in
        for(POIXMLDocumentPart p : getRelations()){
        	//TODO: test to dump p
            if(p instanceof CommentsTable) {
               sheetComments = (CommentsTable)p;
               break;
            }
            if(p instanceof XSSFTable) {
               tables.put( p.getPackageRelationship().getId(), (XSSFTable)p );
            }
        }
        
        // Process external hyperlinks for the sheet, if there are any
        initHyperlinks();
        
        // 20110930, henrichen: for autofilter
        initAutofilter();
        
    }

	/**
     * Initialize worksheet data when creating a new sheet.
     */
    @Override
    protected void onDocumentCreate(){
        worksheet = newSheet();
        initRows(worksheet);
        columnHelper = new ColumnHelper(worksheet);
        hyperlinks = new ArrayList<XSSFHyperlink>();
    }

    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    private void initRows(CTWorksheet worksheet) {
        _rows = new TreeMap<Integer, XSSFRow>();
        tables = new TreeMap<String, XSSFTable>();
        sharedFormulas = new HashMap<Integer, CTCellFormula>();
        arrayFormulas = new ArrayList<CellRangeAddress>();
        for (CTRow row : worksheet.getSheetData().getRowArray()) {
            XSSFRow r = new XSSFRow(row, this);
            _rows.put(r.getRowNum(), r);
        }
    }

    /**
     * Read hyperlink relations, link them with CTHyperlink beans in this worksheet
     * and initialize the internal array of XSSFHyperlink objects
     */
    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    private void initHyperlinks() {
        hyperlinks = new ArrayList<XSSFHyperlink>();

        if(!worksheet.isSetHyperlinks()) return;

        try {
            PackageRelationshipCollection hyperRels =
                getPackagePart().getRelationshipsByType(XSSFRelation.SHEET_HYPERLINKS.getRelation());

            // Turn each one into a XSSFHyperlink
            for(CTHyperlink hyperlink : worksheet.getHyperlinks().getHyperlinkArray()) {
                PackageRelationship hyperRel = null;
                if(hyperlink.getId() != null) {
                    hyperRel = hyperRels.getRelationshipByID(hyperlink.getId());
                }

                hyperlinks.add( new XSSFHyperlink(hyperlink, hyperRel) );
            }
        } catch (InvalidFormatException e){
            throw new POIXMLException(e);
        }
    }

	/**
     * Create a new CTWorksheet instance with all values set to defaults
     *
     * @return a new instance
     */
    private static CTWorksheet newSheet(){
        CTWorksheet worksheet = CTWorksheet.Factory.newInstance();
        CTSheetFormatPr ctFormat = worksheet.addNewSheetFormatPr();
        ctFormat.setDefaultRowHeight(15.0);

        CTSheetView ctView = worksheet.addNewSheetViews().addNewSheetView();
        ctView.setWorkbookViewId(0);

        worksheet.addNewDimension().setRef("A1");

        worksheet.addNewSheetData();

        CTPageMargins ctMargins = worksheet.addNewPageMargins();
        ctMargins.setBottom(0.75);
        ctMargins.setFooter(0.3);
        ctMargins.setHeader(0.3);
        ctMargins.setLeft(0.7);
        ctMargins.setRight(0.7);
        ctMargins.setTop(0.75);

        return worksheet;
    }

    /**
     * Provide access to the CTWorksheet bean holding this sheet's data
     *
     * @return the CTWorksheet bean holding this sheet's data
     */
    @Internal
    public CTWorksheet getCTWorksheet() {
        return this.worksheet;
    }

    public ColumnHelper getColumnHelper() {
        return columnHelper;
    }

    /**
     * Returns the name of this sheet
     *
     * @return the name of this sheet
     */
    public String getSheetName() {
        return sheet.getName();
    }

    /**
     * Adds a merged region of cells (hence those cells form one).
     *
     * @param region (rowfrom/colfrom-rowto/colto) to merge
     * @return index of this region
     */
    public int addMergedRegion(CellRangeAddress region) {
        region.validate(SpreadsheetVersion.EXCEL2007);

        // throw IllegalStateException if the argument CellRangeAddress intersects with
        // a multi-cell array formula defined in this sheet
        validateArrayFormulas(region);

        CTMergeCells ctMergeCells = worksheet.isSetMergeCells() ? worksheet.getMergeCells() : worksheet.addNewMergeCells();
        CTMergeCell ctMergeCell = ctMergeCells.addNewMergeCell();
        ctMergeCell.setRef(region.formatAsString());
        return ctMergeCells.sizeOfMergeCellArray();
    }

    private void validateArrayFormulas(CellRangeAddress region){
        int firstRow = region.getFirstRow();
        int firstColumn = region.getFirstColumn();
        int lastRow = region.getLastRow();
        int lastColumn = region.getLastColumn();
        for (int rowIn = firstRow; rowIn <= lastRow; rowIn++) {
            for (int colIn = firstColumn; colIn <= lastColumn; colIn++) {
                XSSFRow row = getRow(rowIn);
                if (row == null) continue;

                XSSFCell cell = row.getCell(colIn);
                if(cell == null) continue;

                if(cell.isPartOfArrayFormulaGroup()){
                    CellRangeAddress arrayRange = cell.getArrayFormulaRange();
                    if (arrayRange.getNumberOfCells() > 1 &&
                            ( arrayRange.isInRange(region.getFirstRow(), region.getFirstColumn()) ||
                              arrayRange.isInRange(region.getFirstRow(), region.getFirstColumn()))  ){
                        String msg = "The range " + region.formatAsString() + " intersects with a multi-cell array formula. " +
                                "You cannot merge cells of an array.";
                        throw new IllegalStateException(msg);
                    }
                }
            }
        }

    }

    /**
     * Adjusts the column width to fit the contents.
     *
     * This process can be relatively slow on large sheets, so this should
     *  normally only be called once per column, at the end of your
     *  processing.
     *
     * @param column the column index
     */
   public void autoSizeColumn(int column) {
        autoSizeColumn(column, false);
    }

    /**
     * Adjusts the column width to fit the contents.
     * <p>
     * This process can be relatively slow on large sheets, so this should
     *  normally only be called once per column, at the end of your
     *  processing.
     * </p>
     * You can specify whether the content of merged cells should be considered or ignored.
     *  Default is to ignore merged cells.
     *
     * @param column the column index
     * @param useMergedCells whether to use the contents of merged cells when calculating the width of the column
     */
    public void autoSizeColumn(int column, boolean useMergedCells) {
        double width = SheetUtil.getColumnWidth(this, column, useMergedCells);

        if (width != -1) {
            width *= 256;
            int maxColumnWidth = 255*256; // The maximum column width for an individual cell is 255 characters
            if (width > maxColumnWidth) {
                width = maxColumnWidth;
            }
            setColumnWidth(column, (int)(width));
            columnHelper.setColBestFit(column, true);
        }
    }

    /**
     * Create a new SpreadsheetML drawing. If this sheet already contains a drawing - return that.
     *
     * @return a SpreadsheetML drawing
     */
    public XSSFDrawing createDrawingPatriarch() {
        XSSFDrawing drawing = null;
        CTDrawing ctDrawing = getCTDrawing();
        if(ctDrawing == null) {
            //drawingNumber = #drawings.size() + 1
            int drawingNumber = getPackagePart().getPackage().getPartsByContentType(XSSFRelation.DRAWINGS.getContentType()).size() + 1;
            drawing = (XSSFDrawing)createRelationship(XSSFRelation.DRAWINGS, XSSFFactory.getInstance(), drawingNumber);
            String relId = drawing.getPackageRelationship().getId();

            //add CT_Drawing element which indicates that this sheet contains drawing components built on the drawingML platform.
            //The relationship Id references the part containing the drawingML definitions.
            ctDrawing = worksheet.addNewDrawing();
            ctDrawing.setId(relId);
        } else {
            //search the referenced drawing in the list of the sheet's relations
            for(POIXMLDocumentPart p : getRelations()){
                if(p instanceof XSSFDrawing) {
                    XSSFDrawing dr = (XSSFDrawing)p;
                    String drId = dr.getPackageRelationship().getId();
                    if(drId.equals(ctDrawing.getId())){
                        drawing = dr;
                        break;
                    }
                    break;
                }
            }
            if(drawing == null){
                logger.log(POILogger.ERROR, "Can't find drawing with id=" + ctDrawing.getId() + " in the list of the sheet's relationships");
            }
        }
        return drawing;
    }

    /**
     * Get VML drawing for this sheet (aka 'legacy' drawig)
     *
     * @param autoCreate if true, then a new VML drawing part is created
     *
     * @return the VML drawing of <code>null</code> if the drawing was not found and autoCreate=false
     */
    protected XSSFVMLDrawing getVMLDrawing(boolean autoCreate) {
        XSSFVMLDrawing drawing = null;
        CTLegacyDrawing ctDrawing = getCTLegacyDrawing();
        if(ctDrawing == null) {
            if(autoCreate) {
                //drawingNumber = #drawings.size() + 1
                int drawingNumber = getPackagePart().getPackage().getPartsByContentType(XSSFRelation.VML_DRAWINGS.getContentType()).size() + 1;
                drawing = (XSSFVMLDrawing)createRelationship(XSSFRelation.VML_DRAWINGS, XSSFFactory.getInstance(), drawingNumber);
                String relId = drawing.getPackageRelationship().getId();

                //add CTLegacyDrawing element which indicates that this sheet contains drawing components built on the drawingML platform.
                //The relationship Id references the part containing the drawing definitions.
                ctDrawing = worksheet.addNewLegacyDrawing();
                ctDrawing.setId(relId);
            }
        } else {
            //search the referenced drawing in the list of the sheet's relations
            for(POIXMLDocumentPart p : getRelations()){
                if(p instanceof XSSFVMLDrawing) {
                    XSSFVMLDrawing dr = (XSSFVMLDrawing)p;
                    String drId = dr.getPackageRelationship().getId();
                    if(drId.equals(ctDrawing.getId())){
                        drawing = dr;
                        break;
                    }
                    break;
                }
            }
            if(drawing == null){
                logger.log(POILogger.ERROR, "Can't find VML drawing with id=" + ctDrawing.getId() + " in the list of the sheet's relationships");
            }
        }
        return drawing;
    }
    
    protected CTDrawing getCTDrawing() {
       return worksheet.getDrawing();
    }
    protected CTLegacyDrawing getCTLegacyDrawing() {
       return worksheet.getLegacyDrawing();
    }

    /**
     * Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
     * @param colSplit      Horizonatal position of split.
     * @param rowSplit      Vertical position of split.
     */
    public void createFreezePane(int colSplit, int rowSplit) {
        createFreezePane( colSplit, rowSplit, colSplit, rowSplit );
    }

    /**
     * Creates a split (freezepane). Any existing freezepane or split pane is overwritten.
     *
     * <p>
     *     If both colSplit and rowSplit are zero then the existing freeze pane is removed
     * </p>
     *
     * @param colSplit      Horizonatal position of split.
     * @param rowSplit      Vertical position of split.
     * @param leftmostColumn   Left column visible in right pane.
     * @param topRow        Top row visible in bottom pane
     */
    public void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow) {
        CTSheetView ctView = getDefaultSheetView();

        // If both colSplit and rowSplit are zero then the existing freeze pane is removed
        if(colSplit == 0 && rowSplit == 0){
            if(ctView.isSetPane()) ctView.unsetPane();
            ctView.setSelectionArray(null);
            return;
        }

        if (!ctView.isSetPane()) {
            ctView.addNewPane();
        }
        CTPane pane = ctView.getPane();

        if (colSplit > 0) {
           pane.setXSplit(colSplit);
        } else {
           if(pane.isSetXSplit()) pane.unsetXSplit();
        }
        if (rowSplit > 0) {
           pane.setYSplit(rowSplit);
        } else {
           if(pane.isSetYSplit()) pane.unsetYSplit();
        }
        
        pane.setState(STPaneState.FROZEN);
        if (rowSplit == 0) {
            pane.setTopLeftCell(new CellReference(0, leftmostColumn).formatAsString());
            pane.setActivePane(STPane.TOP_RIGHT);
        } else if (colSplit == 0) {
            pane.setTopLeftCell(new CellReference(topRow, 0).formatAsString());
            pane.setActivePane(STPane.BOTTOM_LEFT);
        } else {
            pane.setTopLeftCell(new CellReference(topRow, leftmostColumn).formatAsString());
            pane.setActivePane(STPane.BOTTOM_RIGHT);
        }

        ctView.setSelectionArray(null);
        CTSelection sel = ctView.addNewSelection();
        sel.setPane(pane.getActivePane());
    }

    /**
     * Creates a new comment for this sheet. You still
     *  need to assign it to a cell though
     *
     * @deprecated since Nov 2009 this method is not compatible with the common SS interfaces,
     * use {@link org.zkoss.poi.xssf.usermodel.XSSFDrawing#createCellComment
     *  (org.zkoss.poi.ss.usermodel.ClientAnchor)} instead
     */
    @Deprecated
    public XSSFComment createComment() {
        return createDrawingPatriarch().createCellComment(new XSSFClientAnchor());
    }

    /**
     * Create a new row within the sheet and return the high level representation
     *
     * @param rownum  row number
     * @return High level {@link XSSFRow} object representing a row in the sheet
     * @see #removeRow(org.zkoss.poi.ss.usermodel.Row)
     */
    public XSSFRow createRow(int rownum) {
        CTRow ctRow;
        XSSFRow prev = _rows.get(rownum);
        if(prev != null){
            ctRow = prev.getCTRow();
            ctRow.set(CTRow.Factory.newInstance());
        } else {
        	if(_rows.isEmpty() || rownum > _rows.lastKey()) {
        		// we can append the new row at the end
        		ctRow = worksheet.getSheetData().addNewRow();
        	} else {
        		// get number of rows where row index < rownum
        		// --> this tells us where our row should go
        		int idx = _rows.headMap(rownum).size();
        		ctRow = worksheet.getSheetData().insertNewRow(idx);
        	}
        }
        XSSFRow r = new XSSFRow(ctRow, this);
        r.setRowNum(rownum);
        _rows.put(rownum, r);
        return r;
    }

    /**
     * Creates a split pane. Any existing freezepane or split pane is overwritten.
     * @param xSplitPos      Horizonatal position of split (in 1/20th of a point).
     * @param ySplitPos      Vertical position of split (in 1/20th of a point).
     * @param topRow        Top row visible in bottom pane
     * @param leftmostColumn   Left column visible in right pane.
     * @param activePane    Active pane.  One of: PANE_LOWER_RIGHT,
     *                      PANE_UPPER_RIGHT, PANE_LOWER_LEFT, PANE_UPPER_LEFT
     * @see org.zkoss.poi.ss.usermodel.Sheet#PANE_LOWER_LEFT
     * @see org.zkoss.poi.ss.usermodel.Sheet#PANE_LOWER_RIGHT
     * @see org.zkoss.poi.ss.usermodel.Sheet#PANE_UPPER_LEFT
     * @see org.zkoss.poi.ss.usermodel.Sheet#PANE_UPPER_RIGHT
     */
    public void createSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, int activePane) {
        createFreezePane(xSplitPos, ySplitPos, leftmostColumn, topRow);
        getPane().setState(STPaneState.SPLIT);
        getPane().setActivePane(STPane.Enum.forInt(activePane));
    }

    public XSSFComment getCellComment(int row, int column) {
        if (sheetComments == null) {
            return null;
        }

        String ref = new CellReference(row, column).formatAsString();
        CTComment ctComment = sheetComments.getCTComment(ref);
        if(ctComment == null) return null;

        XSSFVMLDrawing vml = getVMLDrawing(false);
        return new XSSFComment(sheetComments, ctComment,
                vml == null ? null : vml.findCommentShape(row, column));
    }

    public XSSFHyperlink getHyperlink(int row, int column) {
        String ref = new CellReference(row, column).formatAsString();
        for(XSSFHyperlink hyperlink : hyperlinks) {
            if(hyperlink.getCellRef().equals(ref)) {
                return hyperlink;
            }
        }
        return null;
    }

    /**
     * Vertical page break information used for print layout view, page layout view, drawing print breaks
     * in normal view, and for printing the worksheet.
     *
     * @return column indexes of all the vertical page breaks, never <code>null</code>
     */
    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public int[] getColumnBreaks() {
        if (!worksheet.isSetColBreaks() || worksheet.getColBreaks().sizeOfBrkArray() == 0) {
            return new int[0];
        }

        CTBreak[] brkArray = worksheet.getColBreaks().getBrkArray();

        int[] breaks = new int[brkArray.length];
        for (int i = 0 ; i < brkArray.length ; i++) {
            CTBreak brk = brkArray[i];
            breaks[i] = (int)brk.getId() - 1;
        }
        return breaks;
    }

    /**
     * Get the actual column width (in units of 1/256th of a character width )
     *
     * <p>
     * Note, the returned  value is always gerater that {@link #getDefaultColumnWidth()} because the latter does not include margins.
     * Actual column width measured as the number of characters of the maximum digit width of the
     * numbers 0, 1, 2, ..., 9 as rendered in the normal style's font. There are 4 pixels of margin
     * padding (two on each side), plus 1 pixel padding for the gridlines.
     * </p>
     *
     * @param columnIndex - the column to set (0-based)
     * @return width - the width in units of 1/256th of a character width
     */
    public int getColumnWidth(int columnIndex) {
        CTCol col = columnHelper.getColumn(columnIndex, false);
        double width = col == null || !col.isSetWidth() ? getDefaultColumnWidth() : col.getWidth();
        return (int)(width*256);
    }

    /**
     * Get the default column width for the sheet (if the columns do not define their own width) in
     * characters.
     * <p>
     * Note, this value is different from {@link #getColumnWidth(int)}. The latter is always greater and includes
     * 4 pixels of margin padding (two on each side), plus 1 pixel padding for the gridlines.
     * </p>
     * @return column width, default value is 8
     */
    public int getDefaultColumnWidth() {
        CTSheetFormatPr pr = worksheet.getSheetFormatPr();
        return pr == null ? 8 : (int)pr.getBaseColWidth();
    }

    /**
     * Get the default row height for the sheet (if the rows do not define their own height) in
     * twips (1/20 of  a point)
     *
     * @return  default row height
     */
    public short getDefaultRowHeight() {
        return (short)(getDefaultRowHeightInPoints() * 20);
    }

    /**
     * Get the default row height for the sheet measued in point size (if the rows do not define their own height).
     *
     * @return  default row height in points
     */
    public float getDefaultRowHeightInPoints() {
        CTSheetFormatPr pr = worksheet.getSheetFormatPr();
        return (float)(pr == null ? 0 : pr.getDefaultRowHeight());
    }

    private CTSheetFormatPr getSheetTypeSheetFormatPr() {
        return worksheet.isSetSheetFormatPr() ?
               worksheet.getSheetFormatPr() :
               worksheet.addNewSheetFormatPr();
    }

    /**
     * Returns the CellStyle that applies to the given
     *  (0 based) column, or null if no style has been
     *  set for that column
     */
    public CellStyle getColumnStyle(int column) {
        int idx = columnHelper.getColDefaultStyle(column);
        return getWorkbook().getCellStyleAt((short)(idx == -1 ? 0 : idx));
    }

    /**
     * Sets whether the worksheet is displayed from right to left instead of from left to right.
     *
     * @param value true for right to left, false otherwise.
     */
    public void setRightToLeft(boolean value)
    {
       CTSheetView view = getDefaultSheetView();
       view.setRightToLeft(value);
    }

    /**
     * Whether the text is displayed in right-to-left mode in the window
     *
     * @return whether the text is displayed in right-to-left mode in the window
     */
    public boolean isRightToLeft()
    {
       CTSheetView view = getDefaultSheetView();
       return view == null ? false : view.getRightToLeft();
    }

    /**
     * Get whether to display the guts or not,
     * default value is true
     *
     * @return boolean - guts or no guts
     */
    public boolean getDisplayGuts() {
        CTSheetPr sheetPr = getSheetTypeSheetPr();
        CTOutlinePr outlinePr = sheetPr.getOutlinePr() == null ? CTOutlinePr.Factory.newInstance() : sheetPr.getOutlinePr();
        return outlinePr.getShowOutlineSymbols();
    }

    /**
     * Set whether to display the guts or not
     *
     * @param value - guts or no guts
     */
    public void setDisplayGuts(boolean value) {
        CTSheetPr sheetPr = getSheetTypeSheetPr();
        CTOutlinePr outlinePr = sheetPr.getOutlinePr() == null ? sheetPr.addNewOutlinePr() : sheetPr.getOutlinePr();
        outlinePr.setShowOutlineSymbols(value);
    }

    /**
     * Gets the flag indicating whether the window should show 0 (zero) in cells containing zero value.
     * When false, cells with zero value appear blank instead of showing the number zero.
     *
     * @return whether all zero values on the worksheet are displayed
     */
    public boolean isDisplayZeros(){
        CTSheetView view = getDefaultSheetView();
        return view == null ? true : view.getShowZeros();
    }

    /**
     * Set whether the window should show 0 (zero) in cells containing zero value.
     * When false, cells with zero value appear blank instead of showing the number zero.
     *
     * @param value whether to display or hide all zero values on the worksheet
     */
    public void setDisplayZeros(boolean value){
        CTSheetView view = getSheetTypeSheetView();
        view.setShowZeros(value);
    }

    /**
     * Gets the first row on the sheet
     *
     * @return the number of the first logical row on the sheet, zero based
     */
    public int getFirstRowNum() {
        return _rows.size() == 0 ? 0 : _rows.firstKey();
    }

    /**
     * Flag indicating whether the Fit to Page print option is enabled.
     *
     * @return <code>true</code>
     */
    public boolean getFitToPage() {
        CTSheetPr sheetPr = getSheetTypeSheetPr();
        CTPageSetUpPr psSetup = (sheetPr == null || !sheetPr.isSetPageSetUpPr()) ?
                CTPageSetUpPr.Factory.newInstance() : sheetPr.getPageSetUpPr();
        return psSetup.getFitToPage();
    }

    private CTSheetPr getSheetTypeSheetPr() {
        if (worksheet.getSheetPr() == null) {
            worksheet.setSheetPr(CTSheetPr.Factory.newInstance());
        }
        return worksheet.getSheetPr();
    }

    private CTHeaderFooter getSheetTypeHeaderFooter() {
        if (worksheet.getHeaderFooter() == null) {
            worksheet.setHeaderFooter(CTHeaderFooter.Factory.newInstance());
        }
        return worksheet.getHeaderFooter();
    }

    /**
     * Returns the default footer for the sheet,
     *  creating one as needed.
     * You may also want to look at
     *  {@link #getFirstFooter()},
     *  {@link #getOddFooter()} and
     *  {@link #getEvenFooter()}
     */
    public Footer getFooter() {
        // The default footer is an odd footer
        return getOddFooter();
    }

    /**
     * Returns the default header for the sheet,
     *  creating one as needed.
     * You may also want to look at
     *  {@link #getFirstHeader()},
     *  {@link #getOddHeader()} and
     *  {@link #getEvenHeader()}
     */
    public Header getHeader() {
        // The default header is an odd header
        return getOddHeader();
    }

    /**
     * Returns the odd footer. Used on all pages unless
     *  other footers also present, when used on only
     *  odd pages.
     */
    public Footer getOddFooter() {
        return new XSSFOddFooter(getSheetTypeHeaderFooter());
    }
    /**
     * Returns the even footer. Not there by default, but
     *  when set, used on even pages.
     */
    public Footer getEvenFooter() {
        return new XSSFEvenFooter(getSheetTypeHeaderFooter());
    }
    /**
     * Returns the first page footer. Not there by
     *  default, but when set, used on the first page.
     */
    public Footer getFirstFooter() {
        return new XSSFFirstFooter(getSheetTypeHeaderFooter());
    }

    /**
     * Returns the odd header. Used on all pages unless
     *  other headers also present, when used on only
     *  odd pages.
     */
    public Header getOddHeader() {
        return new XSSFOddHeader(getSheetTypeHeaderFooter());
    }
    /**
     * Returns the even header. Not there by default, but
     *  when set, used on even pages.
     */
    public Header getEvenHeader() {
        return new XSSFEvenHeader(getSheetTypeHeaderFooter());
    }
    /**
     * Returns the first page header. Not there by
     *  default, but when set, used on the first page.
     */
    public Header getFirstHeader() {
        return new XSSFFirstHeader(getSheetTypeHeaderFooter());
    }


    /**
     * Determine whether printed output for this sheet will be horizontally centered.
     */
    public boolean getHorizontallyCenter() {
        CTPrintOptions opts = worksheet.getPrintOptions();
        return opts != null && opts.getHorizontalCentered();
    }

    public int getLastRowNum() {
        return _rows.size() == 0 ? 0 : _rows.lastKey();
    }

    public short getLeftCol() {
        String cellRef = worksheet.getSheetViews().getSheetViewArray(0).getTopLeftCell();
        CellReference cellReference = new CellReference(cellRef);
        return cellReference.getCol();
    }

    /**
     * Gets the size of the margin in inches.
     *
     * @param margin which margin to get
     * @return the size of the margin
     * @see Sheet#LeftMargin
     * @see Sheet#RightMargin
     * @see Sheet#TopMargin
     * @see Sheet#BottomMargin
     * @see Sheet#HeaderMargin
     * @see Sheet#FooterMargin
     */
    public double getMargin(short margin) {
        if (!worksheet.isSetPageMargins()) return 0;

        CTPageMargins pageMargins = worksheet.getPageMargins();
        switch (margin) {
            case LeftMargin:
                return pageMargins.getLeft();
            case RightMargin:
                return pageMargins.getRight();
            case TopMargin:
                return pageMargins.getTop();
            case BottomMargin:
                return pageMargins.getBottom();
            case HeaderMargin:
                return pageMargins.getHeader();
            case FooterMargin:
                return pageMargins.getFooter();
            default :
                throw new IllegalArgumentException("Unknown margin constant:  " + margin);
        }
    }

    /**
     * Sets the size of the margin in inches.
     *
     * @param margin which margin to get
     * @param size the size of the margin
     * @see Sheet#LeftMargin
     * @see Sheet#RightMargin
     * @see Sheet#TopMargin
     * @see Sheet#BottomMargin
     * @see Sheet#HeaderMargin
     * @see Sheet#FooterMargin
     */
    public void setMargin(short margin, double size) {
        CTPageMargins pageMargins = worksheet.isSetPageMargins() ?
                worksheet.getPageMargins() : worksheet.addNewPageMargins();
        switch (margin) {
            case LeftMargin:
                pageMargins.setLeft(size);
                break;
            case RightMargin:
                pageMargins.setRight(size);
                break;
            case TopMargin:
                pageMargins.setTop(size);
                break;
            case BottomMargin:
                pageMargins.setBottom(size);
                break;
            case HeaderMargin:
                pageMargins.setHeader(size);
                break;
            case FooterMargin:
                pageMargins.setFooter(size);
                break;
            default :
                throw new IllegalArgumentException( "Unknown margin constant:  " + margin );
        }
    }

    /**
     * @return the merged region at the specified index
     * @throws IllegalStateException if this worksheet does not contain merged regions
     */
    public CellRangeAddress getMergedRegion(int index) {
        CTMergeCells ctMergeCells = worksheet.getMergeCells();
        if(ctMergeCells == null) throw new IllegalStateException("This worksheet does not contain merged regions");

        CTMergeCell ctMergeCell = ctMergeCells.getMergeCellArray(index);
        String ref = ctMergeCell.getRef();
        return CellRangeAddress.valueOf(ref);
    }

    /**
     * Returns the number of merged regions defined in this worksheet
     *
     * @return number of merged regions in this worksheet
     */
    public int getNumMergedRegions() {
        CTMergeCells ctMergeCells = worksheet.getMergeCells();
        return ctMergeCells == null ? 0 : ctMergeCells.sizeOfMergeCellArray();
    }

    public int getNumHyperlinks() {
        return hyperlinks.size();
    }

    /**
     * Returns the information regarding the currently configured pane (split or freeze).
     *
     * @return null if no pane configured, or the pane information.
     */
    public PaneInformation getPaneInformation() {
        CTPane pane = getDefaultSheetView().getPane();
        // no pane configured
        if(pane == null)  return null;

        CellReference cellRef = pane.isSetTopLeftCell() ? new CellReference(pane.getTopLeftCell()) : null;
        return new PaneInformation((short)pane.getXSplit(), (short)pane.getYSplit(),
                (short)(cellRef == null ? 0 : cellRef.getRow()),(cellRef == null ? 0 : cellRef.getCol()),
                (byte)(pane.getActivePane().intValue() - 1), pane.getState() == STPaneState.FROZEN);
    }

    /**
     * Returns the number of phsyically defined rows (NOT the number of rows in the sheet)
     *
     * @return the number of phsyically defined rows
     */
    public int getPhysicalNumberOfRows() {
        return _rows.size();
    }

    /**
     * Gets the print setup object.
     *
     * @return The user model for the print setup object.
     */
    public XSSFPrintSetup getPrintSetup() {
        return new XSSFPrintSetup(worksheet);
    }

    /**
     * Answer whether protection is enabled or disabled
     *
     * @return true => protection enabled; false => protection disabled
     */
    public boolean getProtect() {
        return worksheet.isSetSheetProtection() && sheetProtectionEnabled();
    }
 
    /**
     * Enables sheet protection and sets the password for the sheet.
     * Also sets some attributes on the {@link CTSheetProtection} that correspond to
     * the default values used by Excel
     * 
     * @param password to set for protection. Pass <code>null</code> to remove protection
     */
    public void protectSheet(String password) {
        	
    	if(password != null) {
    		CTSheetProtection sheetProtection = worksheet.addNewSheetProtection();
    		sheetProtection.xsetPassword(stringToExcelPassword(password));
    		sheetProtection.setSheet(true);
    		sheetProtection.setScenarios(true);
    		sheetProtection.setObjects(true);
    	} else {
    		worksheet.unsetSheetProtection();
    	}
    }

	/**
	 * Converts a String to a {@link STUnsignedShortHex} value that contains the {@link PasswordRecord#hashPassword(String)}
	 * value in hexadecimal format
	 *  
	 * @param password the password string you wish convert to an {@link STUnsignedShortHex}
	 * @return {@link STUnsignedShortHex} that contains Excel hashed password in Hex format
	 */
	private STUnsignedShortHex stringToExcelPassword(String password) {
		STUnsignedShortHex hexPassword = STUnsignedShortHex.Factory.newInstance();
		hexPassword.setStringValue(String.valueOf(HexDump.shortToHex(PasswordRecord.hashPassword(password))).substring(2));
		return hexPassword;
	}
	
    /**
     * Returns the logical row ( 0-based).  If you ask for a row that is not
     * defined you get a null.  This is to say row 4 represents the fifth row on a sheet.
     *
     * @param rownum  row to get
     * @return <code>XSSFRow</code> representing the rownumber or <code>null</code> if its not defined on the sheet
     */
    public XSSFRow getRow(int rownum) {
        return _rows.get(rownum);
    }

    /**
     * Horizontal page break information used for print layout view, page layout view, drawing print breaks in normal
     *  view, and for printing the worksheet.
     *
     * @return row indexes of all the horizontal page breaks, never <code>null</code>
     */
    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public int[] getRowBreaks() {
        if (!worksheet.isSetRowBreaks() || worksheet.getRowBreaks().sizeOfBrkArray() == 0) {
            return new int[0];
        }

        CTBreak[] brkArray = worksheet.getRowBreaks().getBrkArray();
        int[] breaks = new int[brkArray.length];
        for (int i = 0 ; i < brkArray.length ; i++) {
            CTBreak brk = brkArray[i];
            breaks[i] = (int)brk.getId() - 1;
        }
        return breaks;
    }

    /**
     * Flag indicating whether summary rows appear below detail in an outline, when applying an outline.
     *
     * <p>
     * When true a summary row is inserted below the detailed data being summarized and a
     * new outline level is established on that row.
     * </p>
     * <p>
     * When false a summary row is inserted above the detailed data being summarized and a new outline level
     * is established on that row.
     * </p>
     * @return <code>true</code> if row summaries appear below detail in the outline
     */
    public boolean getRowSumsBelow() {
        CTSheetPr sheetPr = worksheet.getSheetPr();
        CTOutlinePr outlinePr = (sheetPr != null && sheetPr.isSetOutlinePr())
                ? sheetPr.getOutlinePr() : null;
        return outlinePr == null || outlinePr.getSummaryBelow();
    }

    /**
     * Flag indicating whether summary rows appear below detail in an outline, when applying an outline.
     *
     * <p>
     * When true a summary row is inserted below the detailed data being summarized and a
     * new outline level is established on that row.
     * </p>
     * <p>
     * When false a summary row is inserted above the detailed data being summarized and a new outline level
     * is established on that row.
     * </p>
     * @param value <code>true</code> if row summaries appear below detail in the outline
     */
    public void setRowSumsBelow(boolean value) {
        ensureOutlinePr().setSummaryBelow(value);
    }

    /**
     * Flag indicating whether summary columns appear to the right of detail in an outline, when applying an outline.
     *
     * <p>
     * When true a summary column is inserted to the right of the detailed data being summarized
     * and a new outline level is established on that column.
     * </p>
     * <p>
     * When false a summary column is inserted to the left of the detailed data being
     * summarized and a new outline level is established on that column.
     * </p>
     * @return <code>true</code> if col summaries appear right of the detail in the outline
     */
    public boolean getRowSumsRight() {
        CTSheetPr sheetPr = worksheet.getSheetPr();
        CTOutlinePr outlinePr = (sheetPr != null && sheetPr.isSetOutlinePr())
                ? sheetPr.getOutlinePr() : CTOutlinePr.Factory.newInstance();
        return outlinePr.getSummaryRight();
    }

    /**
     * Flag indicating whether summary columns appear to the right of detail in an outline, when applying an outline.
     *
     * <p>
     * When true a summary column is inserted to the right of the detailed data being summarized
     * and a new outline level is established on that column.
     * </p>
     * <p>
     * When false a summary column is inserted to the left of the detailed data being
     * summarized and a new outline level is established on that column.
     * </p>
     * @param value <code>true</code> if col summaries appear right of the detail in the outline
     */
    public void setRowSumsRight(boolean value) {
        ensureOutlinePr().setSummaryRight(value);
    }


    /**
     * Ensure CTWorksheet.CTSheetPr.CTOutlinePr
     */
    private CTOutlinePr ensureOutlinePr(){
        CTSheetPr sheetPr = worksheet.isSetSheetPr() ? worksheet.getSheetPr() : worksheet.addNewSheetPr();
        return sheetPr.isSetOutlinePr() ? sheetPr.getOutlinePr() : sheetPr.addNewOutlinePr();
    }

    /**
     * A flag indicating whether scenarios are locked when the sheet is protected.
     *
     * @return true => protection enabled; false => protection disabled
     */
    public boolean getScenarioProtect() {
        return worksheet.isSetSheetProtection() && worksheet.getSheetProtection().getScenarios();
    }

    /**
     * The top row in the visible view when the sheet is
     * first viewed after opening it in a viewer
     *
     * @return integer indicating the rownum (0 based) of the top row
     */
    public short getTopRow() {
        String cellRef = getSheetTypeSheetView().getTopLeftCell();
        CellReference cellReference = new CellReference(cellRef);
        return (short) cellReference.getRow();
    }

    /**
     * Determine whether printed output for this sheet will be vertically centered.
     *
     * @return whether printed output for this sheet will be vertically centered.
     */
    public boolean getVerticallyCenter() {
        CTPrintOptions opts = worksheet.getPrintOptions();
        return opts != null && opts.getVerticalCentered();
    }

    /**
     * Group between (0 based) columns
     */
    public void groupColumn(int fromColumn, int toColumn) {
        groupColumn1Based(fromColumn+1, toColumn+1);
    }
    private void groupColumn1Based(int fromColumn, int toColumn) {
        CTCols ctCols=worksheet.getColsArray(0);
        CTCol ctCol=CTCol.Factory.newInstance();
        ctCol.setMin(fromColumn);
        ctCol.setMax(toColumn);
        this.columnHelper.addCleanColIntoCols(ctCols, ctCol);
        for(int index=fromColumn;index<=toColumn;index++){
            CTCol col=columnHelper.getColumn1Based(index, false);
            //col must exist
            short outlineLevel=col.getOutlineLevel();
            col.setOutlineLevel((short)(outlineLevel+1));
            index=(int)col.getMax();
        }
        worksheet.setColsArray(0,ctCols);
        setSheetFormatPrOutlineLevelCol();
    }

    /**
     * Tie a range of cell together so that they can be collapsed or expanded
     *
     * @param fromRow   start row (0-based)
     * @param toRow     end row (0-based)
     */
    public void groupRow(int fromRow, int toRow) {
        for (int i = fromRow; i <= toRow; i++) {
            XSSFRow xrow = getRow(i);
            if (xrow == null) {
                xrow = createRow(i);
            }
            CTRow ctrow = xrow.getCTRow();
            short outlineLevel = ctrow.getOutlineLevel();
            ctrow.setOutlineLevel((short) (outlineLevel + 1));
        }
        setSheetFormatPrOutlineLevelRow();
    }

    private short getMaxOutlineLevelRows(){
        short outlineLevel=0;
        for(XSSFRow xrow : _rows.values()){
            outlineLevel=xrow.getCTRow().getOutlineLevel()>outlineLevel? xrow.getCTRow().getOutlineLevel(): outlineLevel;
        }
        return outlineLevel;
    }


    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    private short getMaxOutlineLevelCols() {
        CTCols ctCols = worksheet.getColsArray(0);
        short outlineLevel = 0;
        for (CTCol col : ctCols.getColArray()) {
            outlineLevel = col.getOutlineLevel() > outlineLevel ? col.getOutlineLevel() : outlineLevel;
        }
        return outlineLevel;
    }

    /**
     * Determines if there is a page break at the indicated column
     */
    public boolean isColumnBroken(int column) {
        int[] colBreaks = getColumnBreaks();
        for (int i = 0 ; i < colBreaks.length ; i++) {
            if (colBreaks[i] == column) {
                return true;
            }
        }
        return false;
    }

    /**
     * Get the hidden state for a given column.
     *
     * @param columnIndex - the column to set (0-based)
     * @return hidden - <code>false</code> if the column is visible
     */
    public boolean isColumnHidden(int columnIndex) {
        CTCol col = columnHelper.getColumn(columnIndex, false);
        return col != null && col.getHidden();
    }

    /**
     * Gets the flag indicating whether this sheet should display formulas.
     *
     * @return <code>true</code> if this sheet should display formulas.
     */
    public boolean isDisplayFormulas() {
        return getSheetTypeSheetView().getShowFormulas();
    }

    /**
     * Gets the flag indicating whether this sheet displays the lines
     * between rows and columns to make editing and reading easier.
     *
     * @return <code>true</code> if this sheet displays gridlines.
     * @see #isPrintGridlines() to check if printing of gridlines is turned on or off
     */
    public boolean isDisplayGridlines() {
        return getSheetTypeSheetView().getShowGridLines();
    }

    /**
     * Sets the flag indicating whether this sheet should display the lines
     * between rows and columns to make editing and reading easier.
     * To turn printing of gridlines use {@link #setPrintGridlines(boolean)}
     *
     *
     * @param show <code>true</code> if this sheet should display gridlines.
     * @see #setPrintGridlines(boolean)
     */
    public void setDisplayGridlines(boolean show) {
        getSheetTypeSheetView().setShowGridLines(show);
    }

    /**
     * Gets the flag indicating whether this sheet should display row and column headings.
     * <p>
     * Row heading are the row numbers to the side of the sheet
     * </p>
     * <p>
     * Column heading are the letters or numbers that appear above the columns of the sheet
     * </p>
     *
     * @return <code>true</code> if this sheet should display row and column headings.
     */
    public boolean isDisplayRowColHeadings() {
        return getSheetTypeSheetView().getShowRowColHeaders();
    }

    /**
     * Sets the flag indicating whether this sheet should display row and column headings.
     * <p>
     * Row heading are the row numbers to the side of the sheet
     * </p>
     * <p>
     * Column heading are the letters or numbers that appear above the columns of the sheet
     * </p>
     *
     * @param show <code>true</code> if this sheet should display row and column headings.
     */
    public void setDisplayRowColHeadings(boolean show) {
        getSheetTypeSheetView().setShowRowColHeaders(show);
    }

    /**
     * Returns whether gridlines are printed.
     *
     * @return whether gridlines are printed
     */
    public boolean isPrintGridlines() {
        CTPrintOptions opts = worksheet.getPrintOptions();
        return opts != null && opts.getGridLines();
    }

    /**
     * Turns on or off the printing of gridlines.
     *
     * @param value boolean to turn on or off the printing of gridlines
     */
    public void setPrintGridlines(boolean value) {
        CTPrintOptions opts = worksheet.isSetPrintOptions() ?
                worksheet.getPrintOptions() : worksheet.addNewPrintOptions();
        opts.setGridLines(value);
    }

    /**
     * Tests if there is a page break at the indicated row
     *
     * @param row index of the row to test
     * @return <code>true</code> if there is a page break at the indicated row
     */
    public boolean isRowBroken(int row) {
        int[] rowBreaks = getRowBreaks();
        for (int i = 0 ; i < rowBreaks.length ; i++) {
            if (rowBreaks[i] == row) {
                return true;
            }
        }
        return false;
    }

    /**
     * Sets a page break at the indicated row
     * Breaks occur above the specified row and left of the specified column inclusive.
     *
     * For example, <code>sheet.setColumnBreak(2);</code> breaks the sheet into two parts
     * with columns A,B,C in the first and D,E,... in the second. Simuilar, <code>sheet.setRowBreak(2);</code>
     * breaks the sheet into two parts with first three rows (rownum=1...3) in the first part
     * and rows starting with rownum=4 in the second.
     *
     * @param row the row to break, inclusive
     */
    public void setRowBreak(int row) {
        CTPageBreak pgBreak = worksheet.isSetRowBreaks() ? worksheet.getRowBreaks() : worksheet.addNewRowBreaks();
        if (! isRowBroken(row)) {
            CTBreak brk = pgBreak.addNewBrk();
            brk.setId(row + 1); // this is id of the row element which is 1-based: <row r="1" ... >
            brk.setMan(true);
            brk.setMax(SpreadsheetVersion.EXCEL2007.getLastColumnIndex()); //end column of the break

            pgBreak.setCount(pgBreak.sizeOfBrkArray());
            pgBreak.setManualBreakCount(pgBreak.sizeOfBrkArray());
        }
    }

    /**
     * Removes a page break at the indicated column
     */
    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public void removeColumnBreak(int column) {
        if (!worksheet.isSetColBreaks()) {
            // no breaks
            return;
        }

        CTPageBreak pgBreak = worksheet.getColBreaks();
        CTBreak[] brkArray = pgBreak.getBrkArray();
        for (int i = 0 ; i < brkArray.length ; i++) {
            if (brkArray[i].getId() == (column + 1)) {
                pgBreak.removeBrk(i);
            }
        }
    }

    /**
     * Removes a merged region of cells (hence letting them free)
     *
     * @param index of the region to unmerge
     */
    public void removeMergedRegion(int index) {
        CTMergeCells ctMergeCells = worksheet.getMergeCells();

        CTMergeCell[] mergeCellsArray = new CTMergeCell[ctMergeCells.sizeOfMergeCellArray() - 1];
        for (int i = 0 ; i < ctMergeCells.sizeOfMergeCellArray() ; i++) {
            if (i < index) {
                mergeCellsArray[i] = ctMergeCells.getMergeCellArray(i);
            }
            else if (i > index) {
                mergeCellsArray[i - 1] = ctMergeCells.getMergeCellArray(i);
            }
        }
        if(mergeCellsArray.length > 0){
            ctMergeCells.setMergeCellArray(mergeCellsArray);
        } else{
            worksheet.unsetMergeCells();
        }
    }

    /**
     * Remove a row from this sheet.  All cells contained in the row are removed as well
     *
     * @param row  the row to remove.
     */
    public void removeRow(Row row) {
        if (row.getSheet() != this) {
            throw new IllegalArgumentException("Specified row does not belong to this sheet");
        }
        // collect cells into a temporary array to avoid ConcurrentModificationException
        ArrayList<XSSFCell> cellsToDelete = new ArrayList<XSSFCell>();
        for(Cell cell : row) cellsToDelete.add((XSSFCell)cell);

        for(XSSFCell cell : cellsToDelete) row.removeCell(cell);

        int idx = _rows.headMap(row.getRowNum()).size();
        _rows.remove(row.getRowNum());
        worksheet.getSheetData().removeRow(idx);
    }

    /**
     * Removes the page break at the indicated row
     */
    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public void removeRowBreak(int row) {
        if(!worksheet.isSetRowBreaks()) {
            return;
        }
        CTPageBreak pgBreak = worksheet.getRowBreaks();
        CTBreak[] brkArray = pgBreak.getBrkArray();
        for (int i = 0 ; i < brkArray.length ; i++) {
            if (brkArray[i].getId() == (row + 1)) {
                pgBreak.removeBrk(i);
            }
        }
    }

    /**
     * Control if Excel should be asked to recalculate all formulas on this sheet
     * when the workbook is opened.
     *
     *  <p>
     *  Calculating the formula values with {@link org.zkoss.poi.ss.usermodel.FormulaEvaluator} is the
     *  recommended solution, but this may be used for certain cases where
     *  evaluation in POI is not possible.
     *  </p>
     *
     *  <p>
     *  It is recommended to force recalcuation of formulas on workbook level using
     *  {@link org.zkoss.poi.ss.usermodel.Workbook#setForceFormulaRecalculation(boolean)}
     *  to ensure that all cross-worksheet formuals and external dependencies are updated.
     *  </p>
     * @param value true if the application will perform a full recalculation of
     * this worksheet values when the workbook is opened
     *
     * @see org.zkoss.poi.ss.usermodel.Workbook#setForceFormulaRecalculation(boolean)
     */
    public void setForceFormulaRecalculation(boolean value) {
       if(worksheet.isSetSheetCalcPr()) {
          // Change the current setting
          CTSheetCalcPr calc = worksheet.getSheetCalcPr();
          calc.setFullCalcOnLoad(value);
       }
       else if(value) {
          // Add the Calc block and set it
          CTSheetCalcPr calc = worksheet.addNewSheetCalcPr();
          calc.setFullCalcOnLoad(value);
       }
       else {
          // Not set, requested not, nothing to do
       }
    }

    /**
     * Whether Excel will be asked to recalculate all formulas when the
     *  workbook is opened.  
     */
    public boolean getForceFormulaRecalculation() {
       if(worksheet.isSetSheetCalcPr()) {
          CTSheetCalcPr calc = worksheet.getSheetCalcPr();
          return calc.getFullCalcOnLoad();
       }
       return false;
    }
    
    /**
     * @return an iterator of the PHYSICAL rows.  Meaning the 3rd element may not
     * be the third row if say for instance the second row is undefined.
     * Call getRowNum() on each row if you care which one it is.
     */
    public Iterator<Row> rowIterator() {
        return (Iterator<Row>)(Iterator<? extends Row>) _rows.values().iterator();
    }

    /**
     * Alias for {@link #rowIterator()} to
     *  allow foreach loops
     */
    public Iterator<Row> iterator() {
        return rowIterator();
    }

    /**
     * Flag indicating whether the sheet displays Automatic Page Breaks.
     *
     * @return <code>true</code> if the sheet displays Automatic Page Breaks.
     */
     public boolean getAutobreaks() {
        CTSheetPr sheetPr = getSheetTypeSheetPr();
        CTPageSetUpPr psSetup = (sheetPr == null || !sheetPr.isSetPageSetUpPr()) ?
                CTPageSetUpPr.Factory.newInstance() : sheetPr.getPageSetUpPr();
        return psSetup.getAutoPageBreaks();
    }

    /**
     * Flag indicating whether the sheet displays Automatic Page Breaks.
     *
     * @param value <code>true</code> if the sheet displays Automatic Page Breaks.
     */
    public void setAutobreaks(boolean value) {
        CTSheetPr sheetPr = getSheetTypeSheetPr();
        CTPageSetUpPr psSetup = sheetPr.isSetPageSetUpPr() ? sheetPr.getPageSetUpPr() : sheetPr.addNewPageSetUpPr();
        psSetup.setAutoPageBreaks(value);
    }

    /**
     * Sets a page break at the indicated column.
     * Breaks occur above the specified row and left of the specified column inclusive.
     *
     * For example, <code>sheet.setColumnBreak(2);</code> breaks the sheet into two parts
     * with columns A,B,C in the first and D,E,... in the second. Simuilar, <code>sheet.setRowBreak(2);</code>
     * breaks the sheet into two parts with first three rows (rownum=1...3) in the first part
     * and rows starting with rownum=4 in the second.
     *
     * @param column the column to break, inclusive
     */
    public void setColumnBreak(int column) {
        if (! isColumnBroken(column)) {
            CTPageBreak pgBreak = worksheet.isSetColBreaks() ? worksheet.getColBreaks() : worksheet.addNewColBreaks();
            CTBreak brk = pgBreak.addNewBrk();
            brk.setId(column + 1);  // this is id of the row element which is 1-based: <row r="1" ... >
            brk.setMan(true);
            brk.setMax(SpreadsheetVersion.EXCEL2007.getLastRowIndex()); //end row of the break

            pgBreak.setCount(pgBreak.sizeOfBrkArray());
            pgBreak.setManualBreakCount(pgBreak.sizeOfBrkArray());
        }
    }

    public void setColumnGroupCollapsed(int columnNumber, boolean collapsed) {
        if (collapsed) {
            collapseColumn(columnNumber);
        } else {
            expandColumn(columnNumber);
        }
    }

    private void collapseColumn(int columnNumber) {
        CTCols cols = worksheet.getColsArray(0);
        CTCol col = columnHelper.getColumn(columnNumber, false);
        int colInfoIx = columnHelper.getIndexOfColumn(cols, col);
        if (colInfoIx == -1) {
            return;
        }
        // Find the start of the group.
        int groupStartColInfoIx = findStartOfColumnOutlineGroup(colInfoIx);

        CTCol columnInfo = cols.getColArray(groupStartColInfoIx);

        // Hide all the columns until the end of the group
        int lastColMax = setGroupHidden(groupStartColInfoIx, columnInfo
                .getOutlineLevel(), true);

        // write collapse field
        setColumn(lastColMax + 1, null, 0, null, null, Boolean.TRUE);

    }

    private void setColumn(int targetColumnIx, Short xfIndex, Integer style,
            Integer level, Boolean hidden, Boolean collapsed) {
        CTCols cols = worksheet.getColsArray(0);
        CTCol ci = null;
        int k = 0;
        for (k = 0; k < cols.sizeOfColArray(); k++) {
            CTCol tci = cols.getColArray(k);
            if (tci.getMin() >= targetColumnIx
                    && tci.getMax() <= targetColumnIx) {
                ci = tci;
                break;
            }
            if (tci.getMin() > targetColumnIx) {
                // call column infos after k are for later columns
                break; // exit now so k will be the correct insert pos
            }
        }

        if (ci == null) {
            // okay so there ISN'T a column info record that covers this column
            // so lets create one!
            CTCol nci = CTCol.Factory.newInstance();
            nci.setMin(targetColumnIx);
            nci.setMax(targetColumnIx);
            unsetCollapsed(collapsed, nci);
            this.columnHelper.addCleanColIntoCols(cols, nci);
            return;
        }

        boolean styleChanged = style != null
        && ci.getStyle() != style;
        boolean levelChanged = level != null
        && ci.getOutlineLevel() != level;
        boolean hiddenChanged = hidden != null
        && ci.getHidden() != hidden;
        boolean collapsedChanged = collapsed != null
        && ci.getCollapsed() != collapsed;
        boolean columnChanged = levelChanged || hiddenChanged
        || collapsedChanged || styleChanged;
        if (!columnChanged) {
            // do nothing...nothing changed.
            return;
        }

        if (ci.getMin() == targetColumnIx && ci.getMax() == targetColumnIx) {
            // ColumnInfo ci for a single column, the target column
            unsetCollapsed(collapsed, ci);
            return;
        }

        if (ci.getMin() == targetColumnIx || ci.getMax() == targetColumnIx) {
            // The target column is at either end of the multi-column ColumnInfo
            // ci
            // we'll just divide the info and create a new one
            if (ci.getMin() == targetColumnIx) {
                ci.setMin(targetColumnIx + 1);
            } else {
                ci.setMax(targetColumnIx - 1);
                k++; // adjust insert pos to insert after
            }
            CTCol nci = columnHelper.cloneCol(cols, ci);
            nci.setMin(targetColumnIx);
            unsetCollapsed(collapsed, nci);
            this.columnHelper.addCleanColIntoCols(cols, nci);

        } else {
            // split to 3 records
            CTCol ciStart = ci;
            CTCol ciMid = columnHelper.cloneCol(cols, ci);
            CTCol ciEnd = columnHelper.cloneCol(cols, ci);
            int lastcolumn = (int) ci.getMax();

            ciStart.setMax(targetColumnIx - 1);

            ciMid.setMin(targetColumnIx);
            ciMid.setMax(targetColumnIx);
            unsetCollapsed(collapsed, ciMid);
            this.columnHelper.addCleanColIntoCols(cols, ciMid);

            ciEnd.setMin(targetColumnIx + 1);
            ciEnd.setMax(lastcolumn);
            this.columnHelper.addCleanColIntoCols(cols, ciEnd);
        }
    }

    private void unsetCollapsed(boolean collapsed, CTCol ci) {
        if (collapsed) {
            ci.setCollapsed(collapsed);
        } else {
            ci.unsetCollapsed();
        }
    }

    /**
     * Sets all adjacent columns of the same outline level to the specified
     * hidden status.
     *
     * @param pIdx
     *                the col info index of the start of the outline group
     * @return the column index of the last column in the outline group
     */
    private int setGroupHidden(int pIdx, int level, boolean hidden) {
        CTCols cols = worksheet.getColsArray(0);
        int idx = pIdx;
        CTCol columnInfo = cols.getColArray(idx);
        while (idx < cols.sizeOfColArray()) {
            columnInfo.setHidden(hidden);
            if (idx + 1 < cols.sizeOfColArray()) {
                CTCol nextColumnInfo = cols.getColArray(idx + 1);

                if (!isAdjacentBefore(columnInfo, nextColumnInfo)) {
                    break;
                }

                if (nextColumnInfo.getOutlineLevel() < level) {
                    break;
                }
                columnInfo = nextColumnInfo;
            }
            idx++;
        }
        return (int) columnInfo.getMax();
    }

    private boolean isAdjacentBefore(CTCol col, CTCol other_col) {
        return (col.getMax() == (other_col.getMin() - 1));
    }

    private int findStartOfColumnOutlineGroup(int pIdx) {
        // Find the start of the group.
        CTCols cols = worksheet.getColsArray(0);
        CTCol columnInfo = cols.getColArray(pIdx);
        int level = columnInfo.getOutlineLevel();
        int idx = pIdx;
        while (idx != 0) {
            CTCol prevColumnInfo = cols.getColArray(idx - 1);
            if (!isAdjacentBefore(prevColumnInfo, columnInfo)) {
                break;
            }
            if (prevColumnInfo.getOutlineLevel() < level) {
                break;
            }
            idx--;
            columnInfo = prevColumnInfo;
        }
        return idx;
    }

    private int findEndOfColumnOutlineGroup(int colInfoIndex) {
        CTCols cols = worksheet.getColsArray(0);
        // Find the end of the group.
        CTCol columnInfo = cols.getColArray(colInfoIndex);
        int level = columnInfo.getOutlineLevel();
        int idx = colInfoIndex;
        while (idx < cols.sizeOfColArray() - 1) {
            CTCol nextColumnInfo = cols.getColArray(idx + 1);
            if (!isAdjacentBefore(columnInfo, nextColumnInfo)) {
                break;
            }
            if (nextColumnInfo.getOutlineLevel() < level) {
                break;
            }
            idx++;
            columnInfo = nextColumnInfo;
        }
        return idx;
    }

    private void expandColumn(int columnIndex) {
        CTCols cols = worksheet.getColsArray(0);
        CTCol col = columnHelper.getColumn(columnIndex, false);
        int colInfoIx = columnHelper.getIndexOfColumn(cols, col);

        int idx = findColInfoIdx((int) col.getMax(), colInfoIx);
        if (idx == -1) {
            return;
        }

        // If it is already expanded do nothing.
        if (!isColumnGroupCollapsed(idx)) {
            return;
        }

        // Find the start/end of the group.
        int startIdx = findStartOfColumnOutlineGroup(idx);
        int endIdx = findEndOfColumnOutlineGroup(idx);

        // expand:
        // colapsed bit must be unset
        // hidden bit gets unset _if_ surrounding groups are expanded you can
        // determine
        // this by looking at the hidden bit of the enclosing group. You will
        // have
        // to look at the start and the end of the current group to determine
        // which
        // is the enclosing group
        // hidden bit only is altered for this outline level. ie. don't
        // uncollapse contained groups
        CTCol columnInfo = cols.getColArray(endIdx);
        if (!isColumnGroupHiddenByParent(idx)) {
            int outlineLevel = columnInfo.getOutlineLevel();
            boolean nestedGroup = false;
            for (int i = startIdx; i <= endIdx; i++) {
                CTCol ci = cols.getColArray(i);
                if (outlineLevel == ci.getOutlineLevel()) {
                    ci.unsetHidden();
                    if (nestedGroup) {
                        nestedGroup = false;
                        ci.setCollapsed(true);
                    }
                } else {
                    nestedGroup = true;
                }
            }
        }
        // Write collapse flag (stored in a single col info record after this
        // outline group)
        setColumn((int) columnInfo.getMax() + 1, null, null, null,
                Boolean.FALSE, Boolean.FALSE);
    }

    private boolean isColumnGroupHiddenByParent(int idx) {
        CTCols cols = worksheet.getColsArray(0);
        // Look out outline details of end
        int endLevel = 0;
        boolean endHidden = false;
        int endOfOutlineGroupIdx = findEndOfColumnOutlineGroup(idx);
        if (endOfOutlineGroupIdx < cols.sizeOfColArray()) {
            CTCol nextInfo = cols.getColArray(endOfOutlineGroupIdx + 1);
            if (isAdjacentBefore(cols.getColArray(endOfOutlineGroupIdx),
                    nextInfo)) {
                endLevel = nextInfo.getOutlineLevel();
                endHidden = nextInfo.getHidden();
            }
        }
        // Look out outline details of start
        int startLevel = 0;
        boolean startHidden = false;
        int startOfOutlineGroupIdx = findStartOfColumnOutlineGroup(idx);
        if (startOfOutlineGroupIdx > 0) {
            CTCol prevInfo = cols.getColArray(startOfOutlineGroupIdx - 1);

            if (isAdjacentBefore(prevInfo, cols
                    .getColArray(startOfOutlineGroupIdx))) {
                startLevel = prevInfo.getOutlineLevel();
                startHidden = prevInfo.getHidden();
            }

        }
        if (endLevel > startLevel) {
            return endHidden;
        }
        return startHidden;
    }

    private int findColInfoIdx(int columnValue, int fromColInfoIdx) {
        CTCols cols = worksheet.getColsArray(0);

        if (columnValue < 0) {
            throw new IllegalArgumentException(
                    "column parameter out of range: " + columnValue);
        }
        if (fromColInfoIdx < 0) {
            throw new IllegalArgumentException(
                    "fromIdx parameter out of range: " + fromColInfoIdx);
        }

        for (int k = fromColInfoIdx; k < cols.sizeOfColArray(); k++) {
            CTCol ci = cols.getColArray(k);

            if (containsColumn(ci, columnValue)) {
                return k;
            }

            if (ci.getMin() > fromColInfoIdx) {
                break;
            }

        }
        return -1;
    }

    private boolean containsColumn(CTCol col, int columnIndex) {
        return col.getMin() <= columnIndex && columnIndex <= col.getMax();
    }

    /**
     * 'Collapsed' state is stored in a single column col info record
     * immediately after the outline group
     *
     * @param idx
     * @return a boolean represented if the column is collapsed
     */
    private boolean isColumnGroupCollapsed(int idx) {
        CTCols cols = worksheet.getColsArray(0);
        int endOfOutlineGroupIdx = findEndOfColumnOutlineGroup(idx);
        int nextColInfoIx = endOfOutlineGroupIdx + 1;
        if (nextColInfoIx >= cols.sizeOfColArray()) {
            return false;
        }
        CTCol nextColInfo = cols.getColArray(nextColInfoIx);

        CTCol col = cols.getColArray(endOfOutlineGroupIdx);
        if (!isAdjacentBefore(col, nextColInfo)) {
            return false;
        }

        return nextColInfo.getCollapsed();
    }

    /**
     * Get the visibility state for a given column.
     *
     * @param columnIndex - the column to get (0-based)
     * @param hidden - the visiblity state of the column
     */
     public void setColumnHidden(int columnIndex, boolean hidden) {
        columnHelper.setColHidden(columnIndex, hidden);
     }

    /**
     * Set the width (in units of 1/256th of a character width)
     *
     * <p>
     * The maximum column width for an individual cell is 255 characters.
     * This value represents the number of characters that can be displayed
     * in a cell that is formatted with the standard font (first font in the workbook).
     * </p>
     *
     * <p>
     * Character width is defined as the maximum digit width
     * of the numbers <code>0, 1, 2, ... 9</code> as rendered
     * using the default font (first font in the workbook).
     * <br/>
     * Unless you are using a very special font, the default character is '0' (zero),
     * this is true for Arial (default font font in HSSF) and Calibri (default font in XSSF)
     * </p>
     *
     * <p>
     * Please note, that the width set by this method includes 4 pixels of margin padding (two on each side),
     * plus 1 pixel padding for the gridlines (Section 3.3.1.12 of the OOXML spec).
     * This results is a slightly less value of visible characters than passed to this method (approx. 1/2 of a character).
     * </p>
     * <p>
     * To compute the actual number of visible characters,
     *  Excel uses the following formula (Section 3.3.1.12 of the OOXML spec):
     * </p>
     * <code>
     *     width = Truncate([{Number of Visible Characters} *
     *      {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
     * </code>
     * <p>Using the Calibri font as an example, the maximum digit width of 11 point font size is 7 pixels (at 96 dpi).
     *  If you set a column width to be eight characters wide, e.g. <code>setColumnWidth(columnIndex, 8*256)</code>,
     *  then the actual value of visible characters (the value shown in Excel) is derived from the following equation:
     *  <code>
            Truncate([numChars*7+5]/7*256)/256 = 8;
     *  </code>
     *
     *  which gives <code>7.29</code>.
     *
     * @param columnIndex - the column to set (0-based)
     * @param width - the width in units of 1/256th of a character width
     * @throws IllegalArgumentException if width > 255*256 (the maximum column width in Excel is 255 characters)
     */
    public void setColumnWidth(int columnIndex, int width) {
        if(width > 255*256) throw new IllegalArgumentException("The maximum column width for an individual cell is 255 characters.");

        columnHelper.setColWidth(columnIndex, (double)width/256);
        columnHelper.setCustomWidth(columnIndex, true);
    }

    public void setDefaultColumnStyle(int column, CellStyle style) {
        columnHelper.setColDefaultStyle(column, style);
    }

    /**
     * Specifies the number of characters of the maximum digit width of the normal style's font.
     * This value does not include margin padding or extra padding for gridlines. It is only the
     * number of characters.
     *
     * @param width the number of characters. Default value is <code>8</code>.
     */
    public void setDefaultColumnWidth(int width) {
        getSheetTypeSheetFormatPr().setBaseColWidth(width);
    }

    /**
     * Set the default row height for the sheet (if the rows do not define their own height) in
     * twips (1/20 of  a point)
     *
     * @param  height default row height in  twips (1/20 of  a point)
     */
    public void setDefaultRowHeight(short height) {
        getSheetTypeSheetFormatPr().setDefaultRowHeight((double)height / 20);

    }

    /**
     * Sets default row height measured in point size.
     *
     * @param height default row height measured in point size.
     */
    public void setDefaultRowHeightInPoints(float height) {
        getSheetTypeSheetFormatPr().setDefaultRowHeight(height);

    }

    /**
     * Sets the flag indicating whether this sheet should display formulas.
     *
     * @param show <code>true</code> if this sheet should display formulas.
     */
    public void setDisplayFormulas(boolean show) {
        getSheetTypeSheetView().setShowFormulas(show);
    }

    private CTSheetView getSheetTypeSheetView() {
        if (getDefaultSheetView() == null) {
            getSheetTypeSheetViews().setSheetViewArray(0, CTSheetView.Factory.newInstance());
        }
        return getDefaultSheetView();
    }

    /**
     * Flag indicating whether the Fit to Page print option is enabled.
     *
     * @param b <code>true</code> if the Fit to Page print option is enabled.
     */
     public void setFitToPage(boolean b) {
        getSheetTypePageSetUpPr().setFitToPage(b);
    }

    /**
     * Center on page horizontally when printing.
     *
     * @param value whether to center on page horizontally when printing.
     */
    public void setHorizontallyCenter(boolean value) {
        CTPrintOptions opts = worksheet.isSetPrintOptions() ?
                worksheet.getPrintOptions() : worksheet.addNewPrintOptions();
        opts.setHorizontalCentered(value);
    }

    /**
     * Whether the output is vertically centered on the page.
     *
     * @param value true to vertically center, false otherwise.
     */
    public void setVerticallyCenter(boolean value) {
        CTPrintOptions opts = worksheet.isSetPrintOptions() ?
                worksheet.getPrintOptions() : worksheet.addNewPrintOptions();
        opts.setVerticalCentered(value);
    }

    /**
     * group the row It is possible for collapsed to be false and yet still have
     * the rows in question hidden. This can be achieved by having a lower
     * outline level collapsed, thus hiding all the child rows. Note that in
     * this case, if the lowest level were expanded, the middle level would
     * remain collapsed.
     *
     * @param rowIndex -
     *                the row involved, 0 based
     * @param collapse -
     *                boolean value for collapse
     */
    public void setRowGroupCollapsed(int rowIndex, boolean collapse) {
        if (collapse) {
            collapseRow(rowIndex);
        } else {
            expandRow(rowIndex);
        }
    }

    /**
     * @param rowIndex the zero based row index to collapse
     */
    private void collapseRow(int rowIndex) {
        XSSFRow row = getRow(rowIndex);
        if (row != null) {
            int startRow = findStartOfRowOutlineGroup(rowIndex);

            // Hide all the columns until the end of the group
            int lastRow = writeHidden(row, startRow, true);
            if (getRow(lastRow) != null) {
                getRow(lastRow).getCTRow().setCollapsed(true);
            } else {
                XSSFRow newRow = createRow(lastRow);
                newRow.getCTRow().setCollapsed(true);
            }
        }
    }

    /**
     * @param rowIndex the zero based row index to find from
     */
    private int findStartOfRowOutlineGroup(int rowIndex) {
        // Find the start of the group.
        int level = getRow(rowIndex).getCTRow().getOutlineLevel();
        int currentRow = rowIndex;
        while (getRow(currentRow) != null) {
            if (getRow(currentRow).getCTRow().getOutlineLevel() < level)
                return currentRow + 1;
            currentRow--;
        }
        return currentRow;
    }

    private int writeHidden(XSSFRow xRow, int rowIndex, boolean hidden) {
        int level = xRow.getCTRow().getOutlineLevel();
        for (Iterator<Row> it = rowIterator(); it.hasNext();) {
            xRow = (XSSFRow) it.next();
            if (xRow.getCTRow().getOutlineLevel() >= level) {
                xRow.getCTRow().setHidden(hidden);
                rowIndex++;
            }

        }
        return rowIndex;
    }

    /**
     * @param rowNumber the zero based row index to expand
     */
    private void expandRow(int rowNumber) {
        if (rowNumber == -1)
            return;
        XSSFRow row = getRow(rowNumber);
        // If it is already expanded do nothing.
        if (!row.getCTRow().isSetHidden())
            return;

        // Find the start of the group.
        int startIdx = findStartOfRowOutlineGroup(rowNumber);

        // Find the end of the group.
        int endIdx = findEndOfRowOutlineGroup(rowNumber);

        // expand:
        // collapsed must be unset
        // hidden bit gets unset _if_ surrounding groups are expanded you can
        // determine
        // this by looking at the hidden bit of the enclosing group. You will
        // have
        // to look at the start and the end of the current group to determine
        // which
        // is the enclosing group
        // hidden bit only is altered for this outline level. ie. don't
        // un-collapse contained groups
        if (!isRowGroupHiddenByParent(rowNumber)) {
            for (int i = startIdx; i < endIdx; i++) {
                if (row.getCTRow().getOutlineLevel() == getRow(i).getCTRow()
                        .getOutlineLevel()) {
                    getRow(i).getCTRow().unsetHidden();
                } else if (!isRowGroupCollapsed(i)) {
                    getRow(i).getCTRow().unsetHidden();
                }
            }
        }
        // Write collapse field
        getRow(endIdx).getCTRow().unsetCollapsed();
    }

    /**
     * @param row the zero based row index to find from
     */
    public int findEndOfRowOutlineGroup(int row) {
        int level = getRow(row).getCTRow().getOutlineLevel();
        int currentRow;
        for (currentRow = row; currentRow < getLastRowNum(); currentRow++) {
            if (getRow(currentRow) == null
                    || getRow(currentRow).getCTRow().getOutlineLevel() < level) {
                break;
            }
        }
        return currentRow;
    }

    /**
     * @param row the zero based row index to find from
     */
    private boolean isRowGroupHiddenByParent(int row) {
        // Look out outline details of end
        int endLevel;
        boolean endHidden;
        int endOfOutlineGroupIdx = findEndOfRowOutlineGroup(row);
        if (getRow(endOfOutlineGroupIdx) == null) {
            endLevel = 0;
            endHidden = false;
        } else {
            endLevel = getRow(endOfOutlineGroupIdx).getCTRow().getOutlineLevel();
            endHidden = getRow(endOfOutlineGroupIdx).getCTRow().getHidden();
        }

        // Look out outline details of start
        int startLevel;
        boolean startHidden;
        int startOfOutlineGroupIdx = findStartOfRowOutlineGroup(row);
        if (startOfOutlineGroupIdx < 0
                || getRow(startOfOutlineGroupIdx) == null) {
            startLevel = 0;
            startHidden = false;
        } else {
            startLevel = getRow(startOfOutlineGroupIdx).getCTRow()
            .getOutlineLevel();
            startHidden = getRow(startOfOutlineGroupIdx).getCTRow()
            .getHidden();
        }
        if (endLevel > startLevel) {
            return endHidden;
        }
        return startHidden;
    }

    /**
     * @param row the zero based row index to find from
     */
    private boolean isRowGroupCollapsed(int row) {
        int collapseRow = findEndOfRowOutlineGroup(row) + 1;
        if (getRow(collapseRow) == null) {
            return false;
        }
        return getRow(collapseRow).getCTRow().getCollapsed();
    }

    /**
     * Sets the zoom magnication for the sheet.  The zoom is expressed as a
     * fraction.  For example to express a zoom of 75% use 3 for the numerator
     * and 4 for the denominator.
     *
     * @param numerator     The numerator for the zoom magnification.
     * @param denominator   The denominator for the zoom magnification.
     * @see #setZoom(int)
     */
    public void setZoom(int numerator, int denominator) {
        int zoom = 100*numerator/denominator;
        setZoom(zoom);
    }

    /**
     * Window zoom magnification for current view representing percent values.
     * Valid values range from 10 to 400. Horizontal & Vertical scale together.
     *
     * For example:
     * <pre>
     * 10 - 10%
     * 20 - 20%
     * ...
     * 100 - 100%
     * ...
     * 400 - 400%
     * </pre>
     *
     * Current view can be Normal, Page Layout, or Page Break Preview.
     *
     * @param scale window zoom magnification
     * @throws IllegalArgumentException if scale is invalid
     */
    public void setZoom(int scale) {
        if(scale < 10 || scale > 400) throw new IllegalArgumentException("Valid scale values range from 10 to 400");
        getSheetTypeSheetView().setZoomScale(scale);
    }

    /**
     * Shifts rows between startRow and endRow n number of rows.
     * If you use a negative number, it will shift rows up.
     * Code ensures that rows don't wrap around.
     *
     * Calls shiftRows(startRow, endRow, n, false, false);
     *
     * <p>
     * Additionally shifts merged regions that are completely defined in these
     * rows (ie. merged 2 cells on a row to be shifted).
     * @param startRow the row to start shifting
     * @param endRow the row to end shifting
     * @param n the number of rows to shift
     */
    public void shiftRows(int startRow, int endRow, int n) {
        shiftRows(startRow, endRow, n, false, false);
    }

    /**
     * Shifts rows between startRow and endRow n number of rows.
     * If you use a negative number, it will shift rows up.
     * Code ensures that rows don't wrap around
     *
     * <p>
     * Additionally shifts merged regions that are completely defined in these
     * rows (ie. merged 2 cells on a row to be shifted).
     * <p>
     * @param startRow the row to start shifting
     * @param endRow the row to end shifting
     * @param n the number of rows to shift
     * @param copyRowHeight whether to copy the row height during the shift
     * @param resetOriginalRowHeight whether to set the original row's height to the default
     */
    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public void shiftRows(int startRow, int endRow, int n, boolean copyRowHeight, boolean resetOriginalRowHeight) {
    	for (Iterator<Row> it = rowIterator() ; it.hasNext() ; ) {
            XSSFRow row = (XSSFRow)it.next();
            int rownum = row.getRowNum();
            if(rownum < startRow) continue;

            if (!copyRowHeight) {
                row.setHeight((short)-1);
            }

            if (removeRow(startRow, endRow, n, rownum)) {
            	// remove row from worksheet.getSheetData row array
            	int idx = _rows.headMap(row.getRowNum()).size();
                worksheet.getSheetData().removeRow(idx);
                // remove row from _rows
                it.remove();
            }
            else if (rownum >= startRow && rownum <= endRow) {
                row.shift(n);
            }

            if(sheetComments != null){
                //TODO shift Note's anchor in the associated /xl/drawing/vmlDrawings#.vml
                CTCommentList lst = sheetComments.getCTComments().getCommentList();
                for (CTComment comment : lst.getCommentArray()) {
                    CellReference ref = new CellReference(comment.getRef());
                    if(ref.getRow() == rownum){
                        ref = new CellReference(rownum + n, ref.getCol());
                        comment.setRef(ref.formatAsString());
                    }
                }
            }
        }
        XSSFRowShifter rowShifter = new XSSFRowShifter(this);

        int sheetIndex = getWorkbook().getSheetIndex(this);
        FormulaShifter shifter = FormulaShifter.createForRowShift(sheetIndex, startRow, endRow, n);

        rowShifter.updateNamedRanges(shifter);
        rowShifter.updateFormulas(shifter);
        rowShifter.shiftMerged(startRow, endRow, n);
        rowShifter.updateConditionalFormatting(shifter);

        //rebuild the _rows map
        TreeMap<Integer, XSSFRow> map = new TreeMap<Integer, XSSFRow>();
        for(XSSFRow r : _rows.values()) {
            map.put(r.getRowNum(), r);
        }
        _rows = map;
    }

    /**
     * Location of the top left visible cell Location of the top left visible cell in the bottom right
     * pane (when in Left-to-Right mode).
     *
     * @param toprow the top row to show in desktop window pane
     * @param leftcol the left column to show in desktop window pane
     */
    public void showInPane(short toprow, short leftcol) {
        CellReference cellReference = new CellReference(toprow, leftcol);
        String cellRef = cellReference.formatAsString();
        getPane().setTopLeftCell(cellRef);
    }

    public void ungroupColumn(int fromColumn, int toColumn) {
        CTCols cols = worksheet.getColsArray(0);
        for (int index = fromColumn; index <= toColumn; index++) {
            CTCol col = columnHelper.getColumn(index, false);
            if (col != null) {
                short outlineLevel = col.getOutlineLevel();
                col.setOutlineLevel((short) (outlineLevel - 1));
                index = (int) col.getMax();

                if (col.getOutlineLevel() <= 0) {
                    int colIndex = columnHelper.getIndexOfColumn(cols, col);
                    worksheet.getColsArray(0).removeCol(colIndex);
                }
            }
        }
        worksheet.setColsArray(0, cols);
        setSheetFormatPrOutlineLevelCol();
    }

    /**
     * Ungroup a range of rows that were previously groupped
     *
     * @param fromRow   start row (0-based)
     * @param toRow     end row (0-based)
     */
    public void ungroupRow(int fromRow, int toRow) {
        for (int i = fromRow; i <= toRow; i++) {
            XSSFRow xrow = getRow(i);
            if (xrow != null) {
                CTRow ctrow = xrow.getCTRow();
                short outlinelevel = ctrow.getOutlineLevel();
                ctrow.setOutlineLevel((short) (outlinelevel - 1));
                //remove a row only if the row has no cell and if the outline level is 0
                if (ctrow.getOutlineLevel() == 0 && xrow.getFirstCellNum() == -1) {
                    removeRow(xrow);
                }
            }
        }
        setSheetFormatPrOutlineLevelRow();
    }

    private void setSheetFormatPrOutlineLevelRow(){
        short maxLevelRow=getMaxOutlineLevelRows();
        getSheetTypeSheetFormatPr().setOutlineLevelRow(maxLevelRow);
    }

    private void setSheetFormatPrOutlineLevelCol(){
        short maxLevelCol=getMaxOutlineLevelCols();
        getSheetTypeSheetFormatPr().setOutlineLevelCol(maxLevelCol);
    }

    private CTSheetViews getSheetTypeSheetViews() {
        if (worksheet.getSheetViews() == null) {
            worksheet.setSheetViews(CTSheetViews.Factory.newInstance());
            worksheet.getSheetViews().addNewSheetView();
        }
        return worksheet.getSheetViews();
    }

    /**
     * Returns a flag indicating whether this sheet is selected.
     * <p>
     * When only 1 sheet is selected and active, this value should be in synch with the activeTab value.
     * In case of a conflict, the Start Part setting wins and sets the active sheet tab.
     * </p>
     * Note: multiple sheets can be selected, but only one sheet can be active at one time.
     *
     * @return <code>true</code> if this sheet is selected
     */
    public boolean isSelected() {
        CTSheetView view = getDefaultSheetView();
        return view != null && view.getTabSelected();
    }

    /**
     * Sets a flag indicating whether this sheet is selected.
     *
     * <p>
     * When only 1 sheet is selected and active, this value should be in synch with the activeTab value.
     * In case of a conflict, the Start Part setting wins and sets the active sheet tab.
     * </p>
     * Note: multiple sheets can be selected, but only one sheet can be active at one time.
     *
     * @param value <code>true</code> if this sheet is selected
     */
    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public void setSelected(boolean value) {
        CTSheetViews views = getSheetTypeSheetViews();
        for (CTSheetView view : views.getSheetViewArray()) {
            view.setTabSelected(value);
        }
    }

    /**
     * Assign a cell comment to a cell region in this worksheet
     *
     * @param cellRef cell region
     * @param comment the comment to assign
     * @deprecated since Nov 2009 use {@link XSSFCell#setCellComment(org.zkoss.poi.ss.usermodel.Comment)} instead
     */
    @Deprecated
    public static void setCellComment(String cellRef, XSSFComment comment) {
        CellReference cellReference = new CellReference(cellRef);

        comment.setRow(cellReference.getRow());
        comment.setColumn(cellReference.getCol());
    }

    /**
     * Register a hyperlink in the collection of hyperlinks on this sheet
     *
     * @param hyperlink the link to add
     */
    @Internal
    public void addHyperlink(XSSFHyperlink hyperlink) {
        hyperlinks.add(hyperlink);
    }

    /**
     * Return location of the active cell, e.g. <code>A1</code>.
     *
     * @return the location of the active cell.
     */
    public String getActiveCell() {
        return getSheetTypeSelection().getActiveCell();
    }

    /**
     * Sets location of the active cell
     *
     * @param cellRef the location of the active cell, e.g. <code>A1</code>..
     */
    public void setActiveCell(String cellRef) {
        CTSelection ctsel = getSheetTypeSelection();
        ctsel.setActiveCell(cellRef);
        ctsel.setSqref(Arrays.asList(cellRef));
    }

    /**
     * Does this sheet have any comments on it? We need to know,
     *  so we can decide about writing it to disk or not
     */
    public boolean hasComments() {
        if(sheetComments == null) { return false; }
        return (sheetComments.getNumberOfComments() > 0);
    }

    protected int getNumberOfComments() {
        if(sheetComments == null) { return 0; }
        return sheetComments.getNumberOfComments();
    }

    private CTSelection getSheetTypeSelection() {
        if (getSheetTypeSheetView().sizeOfSelectionArray() == 0) {
            getSheetTypeSheetView().insertNewSelection(0);
        }
        return getSheetTypeSheetView().getSelectionArray(0);
    }

    /**
     * Return the default sheet view. This is the last one if the sheet's views, according to sec. 3.3.1.83
     * of the OOXML spec: "A single sheet view definition. When more than 1 sheet view is defined in the file,
     * it means that when opening the workbook, each sheet view corresponds to a separate window within the
     * spreadsheet application, where each window is showing the particular sheet. containing the same
     * workbookViewId value, the last sheetView definition is loaded, and the others are discarded.
     * When multiple windows are viewing the same sheet, multiple sheetView elements (with corresponding
     * workbookView entries) are saved."
     */
    private CTSheetView getDefaultSheetView() {
        CTSheetViews views = getSheetTypeSheetViews();
        int sz = views == null ? 0 : views.sizeOfSheetViewArray();
        if (sz  == 0) {
            return null;
        }
        return views.getSheetViewArray(sz - 1);
    }

    /**
     * Returns the sheet's comments object if there is one,
     *  or null if not
     *
     * @param create create a new comments table if it does not exist
     */
    protected CommentsTable getCommentsTable(boolean create) {
        if(sheetComments == null && create){
           // Try to create a comments table with the same number as
           //  the sheet has (i.e. sheet 1 -> comments 1)
           try {
              sheetComments = (CommentsTable)createRelationship(
                    XSSFRelation.SHEET_COMMENTS, XSSFFactory.getInstance(), (int)sheet.getSheetId());
           } catch(PartAlreadyExistsException e) {
              // Technically a sheet doesn't need the same number as
              //  it's comments, and clearly someone has already pinched
              //  our number! Go for the next available one instead
              sheetComments = (CommentsTable)createRelationship(
                    XSSFRelation.SHEET_COMMENTS, XSSFFactory.getInstance(), -1);
           }
        }
        return sheetComments;
    }

    private CTPageSetUpPr getSheetTypePageSetUpPr() {
        CTSheetPr sheetPr = getSheetTypeSheetPr();
        return sheetPr.isSetPageSetUpPr() ? sheetPr.getPageSetUpPr() : sheetPr.addNewPageSetUpPr();
    }

    private boolean removeRow(int startRow, int endRow, int n, int rownum) {
        if (rownum >= (startRow + n) && rownum <= (endRow + n)) {
            if (n > 0 && rownum > endRow) {
                return true;
            }
            else if (n < 0 && rownum < startRow) {
                return true;
            }
        }
        return false;
    }

    private CTPane getPane() {
        if (getDefaultSheetView().getPane() == null) {
            getDefaultSheetView().addNewPane();
        }
        return getDefaultSheetView().getPane();
    }

    /**
     * Return a master shared formula by index
     *
     * @param sid shared group index
     * @return a CTCellFormula bean holding shared formula or <code>null</code> if not found
     */
    CTCellFormula getSharedFormula(int sid){
        return sharedFormulas.get(sid);
    }

    void onReadCell(XSSFCell cell){
        //collect cells holding shared formulas
        CTCell ct = cell.getCTCell();
        CTCellFormula f = ct.getF();
        if (f != null && f.getT() == STCellFormulaType.SHARED && f.isSetRef() && f.getStringValue() != null) {
            // save a detached  copy to avoid XmlValueDisconnectedException,
            // this may happen when the master cell of a shared formula is changed
            sharedFormulas.put((int)f.getSi(), (CTCellFormula)f.copy());
        }
        if (f != null && f.getT() == STCellFormulaType.ARRAY && f.getRef() != null) {
            arrayFormulas.add(CellRangeAddress.valueOf(f.getRef()));
        }
    }

    @Override
    protected void commit() throws IOException {
        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        write(out);
        out.close();
    }

    protected void write(OutputStream out) throws IOException {

        if(worksheet.sizeOfColsArray() == 1) {
            CTCols col = worksheet.getColsArray(0);
            if(col.sizeOfColArray() == 0) {
                worksheet.setColsArray(null);
            }
        }

        // Now re-generate our CTHyperlinks, if needed
        if(hyperlinks.size() > 0) {
            if(worksheet.getHyperlinks() == null) {
                worksheet.addNewHyperlinks();
            }
            CTHyperlink[] ctHls = new CTHyperlink[hyperlinks.size()];
            for(int i=0; i<ctHls.length; i++) {
                // If our sheet has hyperlinks, have them add
                //  any relationships that they might need
                XSSFHyperlink hyperlink = hyperlinks.get(i);
                hyperlink.generateRelationIfNeeded(getPackagePart());
                // Now grab their underling object
                ctHls[i] = hyperlink.getCTHyperlink();
            }
            worksheet.getHyperlinks().setHyperlinkArray(ctHls);
        }

        for(XSSFRow row : _rows.values()){
            row.onDocumentWrite();
        }

        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        xmlOptions.setSaveSyntheticDocumentElement(new QName(CTWorksheet.type.getName().getNamespaceURI(), "worksheet"));
        Map<String, String> map = new HashMap<String, String>();
        map.put(STRelationshipId.type.getName().getNamespaceURI(), "r");
        xmlOptions.setSaveSuggestedPrefixes(map);

        worksheet.save(out, xmlOptions);
    }

    /**
     * @return true when Autofilters are locked and the sheet is protected.
     */
    public boolean isAutoFilterLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getAutoFilter();
    }

    /**
     * @return true when Deleting columns is locked and the sheet is protected.
     */
    public boolean isDeleteColumnsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getDeleteColumns();
    }

    /**
     * @return true when Deleting rows is locked and the sheet is protected.
     */
    public boolean isDeleteRowsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getDeleteRows();
    }

    /**
     * @return true when Formatting cells is locked and the sheet is protected.
     */
    public boolean isFormatCellsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getFormatCells();
    }

    /**
     * @return true when Formatting columns is locked and the sheet is protected.
     */
    public boolean isFormatColumnsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getFormatColumns();
    }

    /**
     * @return true when Formatting rows is locked and the sheet is protected.
     */
    public boolean isFormatRowsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getFormatRows();
    }

    /**
     * @return true when Inserting columns is locked and the sheet is protected.
     */
    public boolean isInsertColumnsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getInsertColumns();
    }

    /**
     * @return true when Inserting hyperlinks is locked and the sheet is protected.
     */
    public boolean isInsertHyperlinksLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getInsertHyperlinks();
    }

    /**
     * @return true when Inserting rows is locked and the sheet is protected.
     */
    public boolean isInsertRowsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getInsertRows();
    }

    /**
     * @return true when Pivot tables are locked and the sheet is protected.
     */
    public boolean isPivotTablesLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getPivotTables();
    }

    /**
     * @return true when Sorting is locked and the sheet is protected.
     */
    public boolean isSortLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getSort();
    }

    /**
     * @return true when Objects are locked and the sheet is protected.
     */
    public boolean isObjectsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getObjects();
    }

    /**
     * @return true when Scenarios are locked and the sheet is protected.
     */
    public boolean isScenariosLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getScenarios();
    }

    /**
     * @return true when Selection of locked cells is locked and the sheet is protected.
     */
    public boolean isSelectLockedCellsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getSelectLockedCells();
    }

    /**
     * @return true when Selection of unlocked cells is locked and the sheet is protected.
     */
    public boolean isSelectUnlockedCellsLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getSelectUnlockedCells();
    }

    /**
     * @return true when Sheet is Protected.
     */
    public boolean isSheetLocked() {
        createProtectionFieldIfNotPresent();
        return sheetProtectionEnabled() && worksheet.getSheetProtection().getSheet();
    }

    /**
     * Enable sheet protection
     */
    public void enableLocking() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setSheet(true);
    }

    /**
     * Disable sheet protection
     */
    public void disableLocking() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setSheet(false);
    }

    /**
     * Enable Autofilters locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockAutoFilter() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setAutoFilter(true);
    }

    /**
     * Enable Deleting columns locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockDeleteColumns() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setDeleteColumns(true);
    }

    /**
     * Enable Deleting rows locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockDeleteRows() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setDeleteRows(true);
    }

    /**
     * Enable Formatting cells locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockFormatCells() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setDeleteColumns(true);
    }

    /**
     * Enable Formatting columns locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockFormatColumns() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setFormatColumns(true);
    }

    /**
     * Enable Formatting rows locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockFormatRows() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setFormatRows(true);
    }

    /**
     * Enable Inserting columns locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockInsertColumns() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setInsertColumns(true);
    }

    /**
     * Enable Inserting hyperlinks locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockInsertHyperlinks() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setInsertHyperlinks(true);
    }

    /**
     * Enable Inserting rows locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockInsertRows() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setInsertRows(true);
    }

    /**
     * Enable Pivot Tables locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockPivotTables() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setPivotTables(true);
    }

    /**
     * Enable Sort locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockSort() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setSort(true);
    }

    /**
     * Enable Objects locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockObjects() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setObjects(true);
    }

    /**
     * Enable Scenarios locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockScenarios() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setScenarios(true);
    }

    /**
     * Enable Selection of locked cells locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockSelectLockedCells() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setSelectLockedCells(true);
    }

    /**
     * Enable Selection of unlocked cells locking.
     * This does not modify sheet protection status.
     * To enforce this locking, call {@link #enableLocking()}
     */
    public void lockSelectUnlockedCells() {
        createProtectionFieldIfNotPresent();
        worksheet.getSheetProtection().setSelectUnlockedCells(true);
    }

    private void createProtectionFieldIfNotPresent() {
        if (worksheet.getSheetProtection() == null) {
            worksheet.setSheetProtection(CTSheetProtection.Factory.newInstance());
        }
    }

    private boolean sheetProtectionEnabled() {
        return worksheet.getSheetProtection().getSheet();
    }

    /* package */ boolean isCellInArrayFormulaContext(XSSFCell cell) {
        for (CellRangeAddress range : arrayFormulas) {
            if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                return true;
            }
        }
        return false;
    }

    /* package */ XSSFCell getFirstCellInArrayFormula(XSSFCell cell) {
        for (CellRangeAddress range : arrayFormulas) {
            if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                return getRow(range.getFirstRow()).getCell(range.getFirstColumn());
            }
        }
        return null;
    }

    /**
     * Also creates cells if they don't exist
     */
    private CellRange<XSSFCell> getCellRange(CellRangeAddress range) {
        int firstRow = range.getFirstRow();
        int firstColumn = range.getFirstColumn();
        int lastRow = range.getLastRow();
        int lastColumn = range.getLastColumn();
        int height = lastRow - firstRow + 1;
        int width = lastColumn - firstColumn + 1;
        List<XSSFCell> temp = new ArrayList<XSSFCell>(height*width);
        for (int rowIn = firstRow; rowIn <= lastRow; rowIn++) {
            for (int colIn = firstColumn; colIn <= lastColumn; colIn++) {
                XSSFRow row = getRow(rowIn);
                if (row == null) {
                    row = createRow(rowIn);
                }
                XSSFCell cell = row.getCell(colIn);
                if (cell == null) {
                    cell = row.createCell(colIn);
                }
                temp.add(cell);
            }
        }
        return SSCellRange.create(firstRow, firstColumn, height, width, temp, XSSFCell.class);
    }

    public CellRange<XSSFCell> setArrayFormula(String formula, CellRangeAddress range) {

        CellRange<XSSFCell> cr = getCellRange(range);

        XSSFCell mainArrayFormulaCell = cr.getTopLeftCell();
        mainArrayFormulaCell.setCellArrayFormula(formula, range);
        arrayFormulas.add(range);
        return cr;
    }

    public CellRange<XSSFCell> removeArrayFormula(Cell cell) {
        if (cell.getSheet() != this) {
            throw new IllegalArgumentException("Specified cell does not belong to this sheet.");
        }
        for (CellRangeAddress range : arrayFormulas) {
            if (range.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                arrayFormulas.remove(range);
                CellRange<XSSFCell> cr = getCellRange(range);
                for (XSSFCell c : cr) {
                    c.setCellType(Cell.CELL_TYPE_BLANK);
                }
                return cr;
            }
        }
        String ref = ((XSSFCell)cell).getCTCell().getR();
        throw new IllegalArgumentException("Cell " + ref + " is not part of an array formula.");
    }


	public DataValidationHelper getDataValidationHelper() {
		return dataValidationHelper;
	}
    
    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    @Override
    public List<DataValidation> getDataValidations() { //20111122, henrichen@zkoss.org: XSSFDataValidation -> DataValidation
    	List<DataValidation> xssfValidations = new ArrayList<DataValidation>(); 
    	CTDataValidations dataValidations = this.worksheet.getDataValidations();
    	if( dataValidations!=null && dataValidations.getCount() > 0 ) {
    		for (CTDataValidation ctDataValidation : dataValidations.getDataValidationArray()) {
    			CellRangeAddressList addressList = new CellRangeAddressList();
    			
    			@SuppressWarnings("unchecked")
    			List<String> sqref = ctDataValidation.getSqref();
    			for (String stRef : sqref) {
    				String[] regions = stRef.split(" ");
    				for (int i = 0; i < regions.length; i++) {
						String[] parts = regions[i].split(":");
						CellReference begin = new CellReference(parts[0]);
						CellReference end = parts.length > 1 ? new CellReference(parts[1]) : begin;
						CellRangeAddress cellRangeAddress = new CellRangeAddress(begin.getRow(), end.getRow(), begin.getCol(), end.getCol());
						addressList.addCellRangeAddress(cellRangeAddress);
					}
				}
				XSSFDataValidation xssfDataValidation = new XSSFDataValidation(addressList, ctDataValidation);
				xssfValidations.add(xssfDataValidation);
			}
    	}
    	return xssfValidations;
    }

	public void addValidationData(DataValidation dataValidation) {
		XSSFDataValidation xssfDataValidation = (XSSFDataValidation)dataValidation;		
		CTDataValidations dataValidations = worksheet.getDataValidations();
		if( dataValidations==null ) {
			dataValidations = worksheet.addNewDataValidations();
		}
		int currentCount = dataValidations.sizeOfDataValidationArray();
        CTDataValidation newval = dataValidations.addNewDataValidation();
		newval.set(xssfDataValidation.getCtDdataValidation());
		dataValidations.setCount(currentCount + 1);

	}
/*
    public XSSFAutoFilter setAutoFilter(CellRangeAddress range) {
        CTAutoFilter af = worksheet.getAutoFilter();
        if(af == null) af = worksheet.addNewAutoFilter();

        CellRangeAddress norm = new CellRangeAddress(range.getFirstRow(), range.getLastRow(),
                range.getFirstColumn(), range.getLastColumn());
        String ref = norm.formatAsString();
        af.setRef(ref);

        XSSFWorkbook wb = getWorkbook();
        int sheetIndex = getWorkbook().getSheetIndex(this);
        XSSFName name = wb.getBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
        if (name == null) {
            name = wb.createBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
            name.getCTName().setHidden(true); 
            CellReference r1 = new CellReference(getSheetName(), range.getFirstRow(), range.getFirstColumn(), true, true);
            CellReference r2 = new CellReference(null, range.getLastRow(), range.getLastColumn(), true, true);
            String fmla = r1.formatAsString() + ":" + r2.formatAsString();
            name.setRefersToFormula(fmla);
        }

        return new XSSFAutoFilter(this);
    }
*/    
     /**
     * Creates a new Table, and associates it with this Sheet
     */
    public XSSFTable createTable() {
       if(! worksheet.isSetTableParts()) {
          worksheet.addNewTableParts();
       }
       
       CTTableParts tblParts = worksheet.getTableParts();
       CTTablePart tbl = tblParts.addNewTablePart();
       
       // Table numbers need to be unique in the file, not just
       //  unique within the sheet. Find the next one
       int tableNumber = getPackagePart().getPackage().getPartsByContentType(XSSFRelation.TABLE.getContentType()).size() + 1;
       XSSFTable table = (XSSFTable)createRelationship(XSSFRelation.TABLE, XSSFFactory.getInstance(), tableNumber);
       tbl.setId(table.getPackageRelationship().getId());
       
       tables.put(tbl.getId(), table);
       
       return table;
    }
    
    /**
     * Returns any tables associated with this Sheet
     */
    public List<XSSFTable> getTables() {
       List<XSSFTable> tableList = new ArrayList<XSSFTable>(
             tables.values()
       );
       return tableList;
    }

    public XSSFSheetConditionalFormatting getSheetConditionalFormatting(){
        return new XSSFSheetConditionalFormatting(this);
    }
	
	//20100914, henrichen@zkoss.org: expose _rows
	protected TreeMap<Integer, XSSFRow> getRows() {
		return _rows;
	}
	//20100914, henrichen@zkoss.org: expose _rows
	protected void setRows(TreeMap<Integer, XSSFRow> rows) {
		_rows = rows;
	}

    //20110506, peterkuo@potix.com
    public boolean isAutoFilterMode() {
        return worksheet.getAutoFilter() != null;
	}

    //TO remove current autofilter
    //20110506, peterkuo@potix.com
    public CellRangeAddress removeAutoFilter(){
    	//remove CTAutoFilter and related name range
    	//TODO:also have to remove related button? send a event?
    	//TODO:also have to restore the height of certain rows
    	CTSheetPr sheetPr = worksheet.getSheetPr();
    	if (sheetPr != null) {
    		sheetPr.setFilterMode(false);
    	}
    	worksheet.unsetAutoFilter();
    	autoFilter = null;
    	
        XSSFWorkbook wb = getWorkbook();
        int sheetIndex = getWorkbook().getSheetIndex(this);
        XSSFName name = wb.getBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
        
    	wb.removeName(name.getNameName());
    	return CellRangeAddress.valueOf(name.getRefersToFormula());
    }

    //20100504: peterkuo@zkoss.org: get autofilter
    private XSSFAutoFilter autoFilter;
    
    @Override
    public AutoFilter getAutoFilter() {
		return autoFilter;
	}
    
    //20110504, henrichen@zkoss.org: set auto filter mode
    /**
     * Set false to remove AutoFilter; set true is ignored.
     * @param b false to remove current AutoFilter; set true is ignored.
     */
    public void setAutoFilterMode(boolean b) {
    	if (!b && isAutoFilterMode()) {
    		removeAutoFilter();
    	}
    }
    
    //20110504, henrichne@zkoss.org: returns whther auto filter mode is in use.
    public boolean isFilterMode() {
    	if (autoFilter != null) {
    		final List<FilterColumn> fcs = autoFilter.getFilterColumns();
    		if (fcs != null) {
    			for(FilterColumn fc : fcs) {
    				List<String> filters = fc.getFilters();
    				if (filters.size() > 0) { //in use
    					return true;
    				}
    			}
    		}
    	}
    	return false;
    }
    
	//20110512, peterkuo@potix.com
	public void removeValidationData(DataValidation dataValidation) {
		XSSFDataValidation xssfDataValidation = (XSSFDataValidation)dataValidation;		
		CTDataValidations dataValidations = worksheet.getDataValidations();
		if( dataValidations==null ) {
			return;
		}
		int currentCount = dataValidations.sizeOfDataValidationArray();
		
		CTDataValidation[] dvArray = dataValidations.getDataValidationArray();
		
		for(int i = 0;i<dvArray.length;i++){
			if(dvArray[i].equals(xssfDataValidation.getCtDdataValidation())){
				dataValidations.removeDataValidation(i);
			}
		}
		dataValidations.setCount(currentCount - 1);
	}

    //20110930, henrichen@zkoss.org, for autofilter
    private void initAutofilter(){
			CTAutoFilter af = worksheet.getAutoFilter();			
			if(af != null){
				autoFilter = new XSSFAutoFilter(this, af);
				fillInAutoFilter(af);
			}
				
    }

    //20110930, henrichen@zkoss.org
    /**
     * Parse from CTAutoFilter to fill the user model autoFilter
     * @param af
     */
    private void fillInAutoFilter(CTAutoFilter af) {
		CTFilterColumn[] fcList = af.getFilterColumnArray();
		if(fcList == null || fcList.length == 0)
			return;
		
		for(CTFilterColumn fc: fcList){
			autoFilter.addFilterColumn(fc);
		}
		
		List<FilterColumn> fcs = autoFilter.getFilterColumns();
		if (fcs != null) {
			for(FilterColumn fc: fcs) {
				((XSSFFilterColumn)fc).init();
			}
		}
	}
    
	//20110930, henrichen@zkoss.org: handle filter
    public AutoFilter setAutoFilter(CellRangeAddress range) {
    	if(isAutoFilterMode()){
    		removeAutoFilter();
    		if (range == null) {
    			return null;
    		}
    	}
    	
        CTAutoFilter ctaf = worksheet.getAutoFilter();
        if(ctaf == null) ctaf = worksheet.addNewAutoFilter();

        String ref = range.formatAsString();
        ctaf.setRef(ref);
        
        final int left = range.getFirstColumn();
        final int top = range.getFirstRow();
        final int right = range.getLastColumn();
        final int bottom = range.getLastRow();
        XSSFWorkbook wb = getWorkbook();
        int sheetIndex = getWorkbook().getSheetIndex(this);
        XSSFName name = wb.getBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
        if (name == null) {
            name = wb.createBuiltInName(XSSFName.BUILTIN_FILTER_DB, sheetIndex);
            name.getCTName().setHidden(true); 
            CellReference r1 = new CellReference(getSheetName(), top, left, true, true);
            CellReference r2 = new CellReference(null, bottom, right, true, true);
            String fmla = r1.formatAsString() + ":" + r2.formatAsString();
            name.setRefersToFormula(fmla);
        }
        
        //set FilterMode in SheetPr to true
        CTSheetPr sheetPr = worksheet.getSheetPr();
        if (sheetPr == null) {
        	sheetPr = worksheet.addNewSheetPr();
        }
        sheetPr.setFilterMode(true);
        
        autoFilter = new XSSFAutoFilter(this, ctaf);
        
        //handle the showButton on merged cell
		for (int i = 0; i < this.getNumMergedRegions(); i++) {
			final CellRangeAddress mrng = this.getMergedRegion(i);
			final int t = mrng.getFirstRow();
	        final int b = mrng.getLastRow();
	        final int l = mrng.getFirstColumn();
	        final int r = mrng.getLastColumn();
	        
	        if (t == top && l <= right && l >= left) { // to be add filter column to hide button
	        	for(int c = l; c < r; ++c) {
		        	final int colId = c - left; 
		        	final XSSFFilterColumn fc = (XSSFFilterColumn) autoFilter.getOrCreateFilterColumn(colId);
		        	fc.setProperties(null, AutoFilter.FILTEROP_AND, null, false);
	        	}
	        }
		}
		
        return autoFilter; 
    }
    
    //20111124, henrichen@zkoss.org:
    public DataValidation getDataValidation(int row, int col) {
    	List<DataValidation> dvs = getDataValidations();
    	if (dvs != null) {
    		for(DataValidation dv : dvs) {
    			CellRangeAddressList addrList = dv.getRegions();
    			for (int j = 0, len = addrList.countRanges(); j < len; ++j) {
    				CellRangeAddress addr = addrList.getCellRangeAddress(j);
    				if (addr.isInRange(row, col)) {
    					return dv;
    				}
    			}
    		}
    	}
    	return null;
    }
}
