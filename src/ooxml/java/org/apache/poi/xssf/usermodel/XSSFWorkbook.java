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

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import javax.xml.namespace.QName;

import org.zkoss.poi.POIXMLDocument;
import org.zkoss.poi.POIXMLDocumentPart;
import org.zkoss.poi.POIXMLException;
import org.zkoss.poi.POIXMLProperties;
import org.zkoss.poi.ss.formula.SheetNameFormatter;
import org.zkoss.poi.openxml4j.exceptions.OpenXML4JException;
import org.zkoss.poi.openxml4j.opc.OPCPackage;
import org.zkoss.poi.openxml4j.opc.PackagePart;
import org.zkoss.poi.openxml4j.opc.PackagePartName;
import org.zkoss.poi.openxml4j.opc.PackageRelationship;
import org.zkoss.poi.openxml4j.opc.PackageRelationshipTypes;
import org.zkoss.poi.openxml4j.opc.PackagingURIHelper;
import org.zkoss.poi.openxml4j.opc.TargetMode;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.usermodel.Row.MissingCellPolicy;
import org.zkoss.poi.ss.util.CellReference;
import org.zkoss.poi.ss.util.WorkbookUtil;
import org.zkoss.poi.util.*;
import org.zkoss.poi.xssf.model.*;
import org.zkoss.poi.xssf.usermodel.helpers.XSSFFormulaUtils;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.officeDocument.x2006.relationships.STRelationshipId;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBookView;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTBookViews;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedNames;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDialogsheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheets;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookPr;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorkbookProtection;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSheetState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.zkoss.poi.xssf.model.CalculationChain;
import org.zkoss.poi.xssf.model.ExternalLink;
import org.zkoss.poi.xssf.model.MapInfo;
import org.zkoss.poi.xssf.model.SharedStringsTable;
import org.zkoss.poi.xssf.model.StylesTable;
import org.zkoss.poi.xssf.model.ThemesTable;

/**
 * High level representation of a SpreadsheetML workbook.  This is the first object most users
 * will construct whether they are reading or writing a workbook.  It is also the
 * top level object for creating new sheets/etc.
 * 
 * @author Henri Chen (henrichen at zkoss dot org) - Sheet1:Sheet3!xxx 3d reference
 */
public class XSSFWorkbook extends POIXMLDocument implements Workbook, Iterable<XSSFSheet> {
    private static final Pattern COMMA_PATTERN = Pattern.compile(",");

    /**
     * Width of one character of the default font in pixels. Same for Calibry and Arial.
     */
    public static final float DEFAULT_CHARACTER_WIDTH = 7.0017f;

    /**
     * Excel silently truncates long sheet names to 31 chars.
     * This constant is used to ensure uniqueness in the first 31 chars
     */
    private static final int MAX_SENSITIVE_SHEET_NAME_LEN = 31;

    /**
     * The underlying XML bean
     */
    private CTWorkbook workbook;

    /**
     * this holds the XSSFSheet objects attached to this workbook
     */
    private List<XSSFSheet> sheets;

    /**
     * this holds the XSSFName objects attached to this workbook
     */
    private List<XSSFName> namedRanges;

    /**
     * shared string table - a cache of strings in this workbook
     */
    private SharedStringsTable sharedStringSource;

    /**
     * A collection of shared objects used for styling content,
     * e.g. fonts, cell styles, colors, etc.
     */
    private StylesTable stylesSource;

    private ThemesTable theme;

    /**
     * The locator of user-defined functions.
     * By default includes functions from the Excel Analysis Toolpack
     */
    private IndexedUDFFinder _udfFinder = new IndexedUDFFinder(UDFFinder.DEFAULT);

    /**
     * TODO
     */
    private CalculationChain calcChain;

    /**
     * A collection of custom XML mappings
     */
    private MapInfo mapInfo;

    /**
     * Used to keep track of the data formatter so that all
     * createDataFormatter calls return the same one for a given
     * book.  This ensures that updates from one places is visible
     * someplace else.
     */
    private XSSFDataFormat formatter;

    /**
     * The policy to apply in the event of missing or
     *  blank cells when fetching from a row.
     * See {@link org.zkoss.poi.ss.usermodel.Row.MissingCellPolicy}
     */
    private MissingCellPolicy _missingCellPolicy = Row.RETURN_NULL_AND_BLANK;

    /**
     * array of pictures for this workbook
     */
    private List<XSSFPictureData> pictures;

    private static POILogger logger = POILogFactory.getLogger(XSSFWorkbook.class);

    /**
     * cached instance of XSSFCreationHelper for this workbook
     * @see {@link #getCreationHelper()}
     */
    private XSSFCreationHelper _creationHelper;

    /**
     * List of Array of external sheet references.
     * String[0]: book name.
     * String[1]: first sheet name
     * String[2]: last sheet name
     */
	private List<String[]> _externalSheetRefs = new ArrayList<String[]>(4);

	/**
	 * Map from external link index to target book name (for evaluation)
	 */
	private Map<String, String> linkIndexToBookName = new HashMap<String, String>(4);
	
	/**
	 * Map from target book name to external link index (for formula parsing)
	 */
	private Map<String, String> bookNameToLinkIndex = new HashMap<String, String>(4);
	
	/**
	 * @return  the external sheet index with the given book name and 
	 * sheet names; create if not exists.
	 */
	/*package*/ int getOrCreateExternalSheetIndex(String bookName, String sheetName1, String sheetName2) {
		synchronized(_externalSheetRefs) {
			final int len = _externalSheetRefs.size();
			for(int j = 0; j < len; ++j) {
				final String jbookName = _externalSheetRefs.get(j)[0];
				if ((bookName == jbookName || (bookName != null && bookName.equalsIgnoreCase(jbookName))) 
					&& _externalSheetRefs.get(j)[1].equalsIgnoreCase(sheetName1) 
					&& _externalSheetRefs.get(j)[2].equalsIgnoreCase(sheetName2)) {
					return j;
				}
			}
			_externalSheetRefs.add(new String[] {bookName, sheetName1, sheetName2});
			return len;
		}
	}
	
	/*package*/ String[] convertFromExternSheetIndex(int externSheetIndex) {
		if (_externalSheetRefs.size() <= externSheetIndex) {
			return null;
		}
		return _externalSheetRefs.get(externSheetIndex);
	}

    /**
     * Create a new SpreadsheetML workbook.
     */
    public XSSFWorkbook() {
        super(newPackage());
        onWorkbookCreate();
    }

    /**
     * Constructs a XSSFWorkbook object given a OpenXML4J <code>Package</code> object,
     * see <a href="http://openxml4j.org/">www.openxml4j.org</a>.
     *
     * @param pkg the OpenXML4J <code>Package</code> object.
     */
    public XSSFWorkbook(OPCPackage pkg) throws IOException {
        super(pkg);

        //build a tree of POIXMLDocumentParts, this workbook being the root
        load(XSSFFactory.getInstance());
    }

    public XSSFWorkbook(InputStream is) throws IOException {
        super(PackageHelper.open(is));
        
        //build a tree of POIXMLDocumentParts, this workbook being the root
        load(XSSFFactory.getInstance());
    }

    /**
     * Constructs a XSSFWorkbook object given a file name.
     *
     * @param      path   the file name.
     */
    public XSSFWorkbook(String path) throws IOException {
        this(openPackage(path));
    }

    @Override
    @SuppressWarnings("deprecation") //  getXYZArray() array accessors are deprecated
    protected void onDocumentRead() throws IOException {
        try {
            WorkbookDocument doc = WorkbookDocument.Factory.parse(getPackagePart().getInputStream());
            this.workbook = doc.getWorkbook();

            Map<String, XSSFSheet> shIdMap = new HashMap<String, XSSFSheet>();
            for(POIXMLDocumentPart p : getRelations()){
                if(p instanceof SharedStringsTable) sharedStringSource = (SharedStringsTable)p;
                else if(p instanceof StylesTable) stylesSource = (StylesTable)p;
                else if(p instanceof ThemesTable) theme = (ThemesTable)p;
                else if(p instanceof CalculationChain) calcChain = (CalculationChain)p;
                else if(p instanceof MapInfo) mapInfo = (MapInfo)p;
                else if (p instanceof XSSFSheet) {
                    shIdMap.put(p.getPackageRelationship().getId(), (XSSFSheet)p);
                } else if (p instanceof ExternalLink) {
                	final ExternalLink el = (ExternalLink) p;
                	linkIndexToBookName.put(el.getLinkIndex(), el.getBookName());
                	bookNameToLinkIndex.put(el.getBookName(), el.getLinkIndex());
                }
            }
            stylesSource.setTheme(theme);

            if(sharedStringSource == null) {
                //Create SST if it is missing
                sharedStringSource = (SharedStringsTable)createRelationship(XSSFRelation.SHARED_STRINGS, XSSFFactory.getInstance());
            }

            // Load individual sheets. The order of sheets is defined by the order of CTSheet elements in the workbook
            sheets = new ArrayList<XSSFSheet>(shIdMap.size());
            for (CTSheet ctSheet : this.workbook.getSheets().getSheetArray()) {
                XSSFSheet sh = shIdMap.get(ctSheet.getId());
                if(sh == null) {
                    logger.log(POILogger.WARN, "Sheet with name " + ctSheet.getName() + " and r:id " + ctSheet.getId()+ " was defined, but didn't exist in package, skipping");
                    continue;
                }
                sh.sheet = ctSheet;
                sh.onDocumentRead();
                sheets.add(sh);
            }

            // Process the named ranges
            namedRanges = new ArrayList<XSSFName>();
            if(workbook.isSetDefinedNames()) {
                for(CTDefinedName ctName : workbook.getDefinedNames().getDefinedNameArray()) {
                    namedRanges.add(new XSSFName(ctName, this));
                }
            }

        } catch (XmlException e) {
            throw new POIXMLException(e);
        }
    }

    /**
     * Create a new CTWorkbook with all values set to default
     */
    private void onWorkbookCreate() {
        workbook = CTWorkbook.Factory.newInstance();

        // don't EVER use the 1904 date system
        CTWorkbookPr workbookPr = workbook.addNewWorkbookPr();
        workbookPr.setDate1904(false);

        CTBookViews bvs = workbook.addNewBookViews();
        CTBookView bv = bvs.addNewWorkbookView();
        bv.setActiveTab(0);
        workbook.addNewSheets();

        POIXMLProperties.ExtendedProperties expProps = getProperties().getExtendedProperties();
        expProps.getUnderlyingProperties().setApplication(DOCUMENT_CREATOR);

        sharedStringSource = (SharedStringsTable)createRelationship(XSSFRelation.SHARED_STRINGS, XSSFFactory.getInstance());
        stylesSource = (StylesTable)createRelationship(XSSFRelation.STYLES, XSSFFactory.getInstance());

        namedRanges = new ArrayList<XSSFName>();
        sheets = new ArrayList<XSSFSheet>();
    }

    /**
     * Create a new SpreadsheetML package and setup the default minimal content
     */
    protected static OPCPackage newPackage() {
        try {
            OPCPackage pkg = OPCPackage.create(new ByteArrayOutputStream());
            // Main part
            PackagePartName corePartName = PackagingURIHelper.createPartName(XSSFRelation.WORKBOOK.getDefaultFileName());
            // Create main part relationship
            pkg.addRelationship(corePartName, TargetMode.INTERNAL, PackageRelationshipTypes.CORE_DOCUMENT);
            // Create main document part
            pkg.createPart(corePartName, XSSFRelation.WORKBOOK.getContentType());

            pkg.getPackageProperties().setCreatorProperty(DOCUMENT_CREATOR);

            return pkg;
        } catch (Exception e){
            throw new POIXMLException(e);
        }
    }

    /**
     * Return the underlying XML bean
     *
     * @return the underlying CTWorkbook bean
     */
    @Internal
    public CTWorkbook getCTWorkbook() {
        return this.workbook;
    }

    /**
     * Adds a picture to the workbook.
     *
     * @param pictureData       The bytes of the picture
     * @param format            The format of the picture.
     *
     * @return the index to this picture (0 based), the added picture can be obtained from {@link #getAllPictures()} .
     * @see Workbook#PICTURE_TYPE_EMF
     * @see Workbook#PICTURE_TYPE_WMF
     * @see Workbook#PICTURE_TYPE_PICT
     * @see Workbook#PICTURE_TYPE_JPEG
     * @see Workbook#PICTURE_TYPE_PNG
     * @see Workbook#PICTURE_TYPE_DIB
     * @see #getAllPictures()
     */
    public int addPicture(byte[] pictureData, int format) {
        int imageNumber = getAllPictures().size() + 1;
        XSSFPictureData img = (XSSFPictureData)createRelationship(XSSFPictureData.RELATIONS[format], XSSFFactory.getInstance(), imageNumber, true);
        try {
            OutputStream out = img.getPackagePart().getOutputStream();
            out.write(pictureData);
            out.close();
        } catch (IOException e){
            throw new POIXMLException(e);
        }
        pictures.add(img);
        return imageNumber - 1;
    }

    /**
     * Adds a picture to the workbook.
     *
     * @param is                The sream to read image from
     * @param format            The format of the picture.
     *
     * @return the index to this picture (0 based), the added picture can be obtained from {@link #getAllPictures()} .
     * @see Workbook#PICTURE_TYPE_EMF
     * @see Workbook#PICTURE_TYPE_WMF
     * @see Workbook#PICTURE_TYPE_PICT
     * @see Workbook#PICTURE_TYPE_JPEG
     * @see Workbook#PICTURE_TYPE_PNG
     * @see Workbook#PICTURE_TYPE_DIB
     * @see #getAllPictures()
     */
    public int addPicture(InputStream is, int format) throws IOException {
        int imageNumber = getAllPictures().size() + 1;
        XSSFPictureData img = (XSSFPictureData)createRelationship(XSSFPictureData.RELATIONS[format], XSSFFactory.getInstance(), imageNumber, true);
        OutputStream out = img.getPackagePart().getOutputStream();
        IOUtils.copy(is, out);
        out.close();
        pictures.add(img);
        return imageNumber - 1;
    }

    /**
     * Create an XSSFSheet from an existing sheet in the XSSFWorkbook.
     *  The cloned sheet is a deep copy of the original.
     *
     * @return XSSFSheet representing the cloned sheet.
     * @throws IllegalArgumentException if the sheet index in invalid
     * @throws POIXMLException if there were errors when cloning
     */
    public XSSFSheet cloneSheet(int sheetNum) {
        validateSheetIndex(sheetNum);

        XSSFSheet srcSheet = sheets.get(sheetNum);
        String srcName = srcSheet.getSheetName();
        String clonedName = getUniqueSheetName(srcName);

        XSSFSheet clonedSheet = createSheet(clonedName);
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            srcSheet.write(out);
            clonedSheet.read(new ByteArrayInputStream(out.toByteArray()));
        } catch (IOException e){
            throw new POIXMLException("Failed to clone sheet", e);
        }
        CTWorksheet ct = clonedSheet.getCTWorksheet();
        if(ct.isSetDrawing()) {
            logger.log(POILogger.WARN, "Cloning sheets with drawings is not yet supported.");
            ct.unsetDrawing();
        }
        if(ct.isSetLegacyDrawing()) {
            logger.log(POILogger.WARN, "Cloning sheets with comments is not yet supported.");
            ct.unsetLegacyDrawing();
        }

        clonedSheet.setSelected(false);
        return clonedSheet;
    }

    /**
     * Generate a valid sheet name based on the existing one. Used when cloning sheets.
     *
     * @param srcName the original sheet name to
     * @return clone sheet name
     */
    private String getUniqueSheetName(String srcName) {
        int uniqueIndex = 2;
        String baseName = srcName;
        int bracketPos = srcName.lastIndexOf('(');
        if (bracketPos > 0 && srcName.endsWith(")")) {
            String suffix = srcName.substring(bracketPos + 1, srcName.length() - ")".length());
            try {
                uniqueIndex = Integer.parseInt(suffix.trim());
                uniqueIndex++;
                baseName = srcName.substring(0, bracketPos).trim();
            } catch (NumberFormatException e) {
                // contents of brackets not numeric
            }
        }
        while (true) {
            // Try and find the next sheet name that is unique
            String index = Integer.toString(uniqueIndex++);
            String name;
            if (baseName.length() + index.length() + 2 < 31) {
                name = baseName + " (" + index + ")";
            } else {
                name = baseName.substring(0, 31 - index.length() - 2) + "(" + index + ")";
            }

            //If the sheet name is unique, then set it otherwise move on to the next number.
            if (getSheetIndex(name) == -1) {
                return name;
            }
        }
    }

    /**
     * Create a new XSSFCellStyle and add it to the workbook's style table
     *
     * @return the new XSSFCellStyle object
     */
    public XSSFCellStyle createCellStyle() {
        return stylesSource.createCellStyle();
    }

    /**
     * Returns the instance of XSSFDataFormat for this workbook.
     *
     * @return the XSSFDataFormat object
     * @see org.zkoss.poi.ss.usermodel.DataFormat
     */
    public XSSFDataFormat createDataFormat() {
        if (formatter == null)
            formatter = new XSSFDataFormat(stylesSource);
        return formatter;
    }

    /**
     * Create a new Font and add it to the workbook's font table
     *
     * @return new font object
     */
    public XSSFFont createFont() {
        XSSFFont font = new XSSFFont();
        font.registerTo(stylesSource);
        return font;
    }

    public XSSFName createName() {
        CTDefinedName ctName = CTDefinedName.Factory.newInstance();
        ctName.setName("");
        XSSFName name = new XSSFName(ctName, this);
        namedRanges.add(name);
        return name;
    }

    /**
     * Create an XSSFSheet for this workbook, adds it to the sheets and returns
     * the high level representation.  Use this to create new sheets.
     *
     * @return XSSFSheet representing the new sheet.
     */
    public XSSFSheet createSheet() {
        String sheetname = "Sheet" + (sheets.size());
        int idx = 0;
        while(getSheet(sheetname) != null) {
            sheetname = "Sheet" + idx;
            idx++;
        }
        return createSheet(sheetname);
    }

    /**
     * Create an XSSFSheet for this workbook, adds it to the sheets and returns
     * the high level representation.  Use this to create new sheets.
     *
     * @param sheetname  sheetname to set for the sheet, can't be duplicate, greater than 31 chars or contain /\?*[]
     * @return XSSFSheet representing the new sheet.
     * @throws IllegalArgumentException if the sheetname is invalid or the workbook already contains a sheet of this name
     */
    public XSSFSheet createSheet(String sheetname) {
        if (containsSheet( sheetname, sheets.size() ))
               throw new IllegalArgumentException( "The workbook already contains a sheet of this name");

        CTSheet sheet = addSheet(sheetname);

        int sheetNumber = 1;
        for(XSSFSheet sh : sheets) sheetNumber = (int)Math.max(sh.sheet.getSheetId() + 1, sheetNumber);

        XSSFSheet wrapper = (XSSFSheet)createRelationship(XSSFRelation.WORKSHEET, XSSFFactory.getInstance(), sheetNumber);
        wrapper.sheet = sheet;
        sheet.setId(wrapper.getPackageRelationship().getId());
        sheet.setSheetId(sheetNumber);
        if(sheets.size() == 0) wrapper.setSelected(true);
        sheets.add(wrapper);
        return wrapper;
    }

    protected XSSFDialogsheet createDialogsheet(String sheetname, CTDialogsheet dialogsheet) {
        XSSFSheet sheet = createSheet(sheetname);
        return new XSSFDialogsheet(sheet);
    }

    private CTSheet addSheet(String sheetname) {
        WorkbookUtil.validateSheetName(sheetname);

        CTSheet sheet = workbook.getSheets().addNewSheet();
        sheet.setName(sheetname);
        return sheet;
    }

    /**
     * Finds a font that matches the one with the supplied attributes
     */
    public XSSFFont findFont(short boldWeight, short color, short fontHeight, String name, boolean italic, boolean strikeout, short typeOffset, byte underline) {
        return stylesSource.findFont(boldWeight, color, fontHeight, name, italic, strikeout, typeOffset, underline);
    }

    /**
     * Convenience method to get the active sheet.  The active sheet is is the sheet
     * which is currently displayed when the workbook is viewed in Excel.
     * 'Selected' sheet(s) is a distinct concept.
     */
    public int getActiveSheetIndex() {
        //activeTab (Active Sheet Index) Specifies an unsignedInt
        //that contains the index to the active sheet in this book view.
        return (int)workbook.getBookViews().getWorkbookViewArray(0).getActiveTab();
    }

    /**
     * Gets all pictures from the Workbook.
     *
     * @return the list of pictures (a list of {@link XSSFPictureData} objects.)
     * @see #addPicture(byte[], int)
     */
    public List<XSSFPictureData> getAllPictures() {
        if(pictures == null) {
            //In OOXML pictures are referred to in sheets,
            //dive into sheet's relations, select drawings and their images
            pictures = new ArrayList<XSSFPictureData>();
            for(XSSFSheet sh : sheets){
                for(POIXMLDocumentPart dr : sh.getRelations()){
                    if(dr instanceof XSSFDrawing){
                        for(POIXMLDocumentPart img : dr.getRelations()){
                            if(img instanceof XSSFPictureData){
                                pictures.add((XSSFPictureData)img);
                            }
                        }
                    }
                }
            }
        }
        return pictures;
    }

    /**
     * gGet the cell style object at the given index
     *
     * @param idx  index within the set of styles
     * @return XSSFCellStyle object at the index
     */
    public XSSFCellStyle getCellStyleAt(short idx) {
        return stylesSource.getStyleAt(idx);
    }

    /**
     * Get the font at the given index number
     *
     * @param idx  index number
     * @return XSSFFont at the index
     */
    public XSSFFont getFontAt(short idx) {
        return stylesSource.getFontAt(idx);
    }

    public XSSFName getName(String name) {
        int nameIndex = getNameIndex(name);
        if (nameIndex < 0) {
            return null;
        }
        return namedRanges.get(nameIndex);
    }

    public XSSFName getNameAt(int nameIndex) {
        int nNames = namedRanges.size();
        if (nNames < 1) {
            throw new IllegalStateException("There are no defined names in this workbook");
        }
        if (nameIndex < 0 || nameIndex > nNames) {
            throw new IllegalArgumentException("Specified name index " + nameIndex
                    + " is outside the allowable range (0.." + (nNames-1) + ").");
        }
        return namedRanges.get(nameIndex);
    }

    /**
     * Gets the named range index by his name
     * <i>Note:</i>Excel named ranges are case-insensitive and
     * this method performs a case-insensitive search.
     *
     * @param name named range name
     * @return named range index
     */
    public int getNameIndex(String name) {
        int i = 0;
        for(XSSFName nr : namedRanges) {
            if(nr.getNameName().equals(name)) {
                return i;
            }
            i++;
        }
        return -1;
    }

    /**
     * Get the number of styles the workbook contains
     *
     * @return count of cell styles
     */
    public short getNumCellStyles() {
        return (short) (stylesSource).getNumCellStyles();
    }

    /**
     * Get the number of fonts in the this workbook
     *
     * @return number of fonts
     */
    public short getNumberOfFonts() {
        return (short)stylesSource.getFonts().size();
    }

    /**
     * Get the number of named ranges in the this workbook
     *
     * @return number of named ranges
     */
    public int getNumberOfNames() {
        return namedRanges.size();
    }

    /**
     * Get the number of worksheets in the this workbook
     *
     * @return number of worksheets
     */
    public int getNumberOfSheets() {
        return sheets.size();
    }

    /**
     * Retrieves the reference for the printarea of the specified sheet, the sheet name is appended to the reference even if it was not specified.
     * @param sheetIndex Zero-based sheet index (0 Represents the first sheet to keep consistent with java)
     * @return String Null if no print area has been defined
     */
    public String getPrintArea(int sheetIndex) {
        XSSFName name = getBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
        if (name == null) return null;
        //adding one here because 0 indicates a global named region; doesnt make sense for print areas
        return name.getRefersToFormula();

    }

    /**
     * Get sheet with the given name (case insensitive match)
     *
     * @param name of the sheet
     * @return XSSFSheet with the name provided or <code>null</code> if it does not exist
     */
    public XSSFSheet getSheet(String name) {
        for (XSSFSheet sheet : sheets) {
            if (name.equalsIgnoreCase(sheet.getSheetName())) {
                return sheet;
            }
        }
        return null;
    }

    /**
     * Get the XSSFSheet object at the given index.
     *
     * @param index of the sheet number (0-based physical & logical)
     * @return XSSFSheet at the provided index
     * @throws IllegalArgumentException if the index is out of range (index
     *            &lt; 0 || index &gt;= getNumberOfSheets()).
     */
    public XSSFSheet getSheetAt(int index) {
        validateSheetIndex(index);
        return sheets.get(index);
    }

    /**
     * Returns the index of the sheet by his name (case insensitive match)
     *
     * @param name the sheet name
     * @return index of the sheet (0 based) or <tt>-1</tt if not found
     */
    public int getSheetIndex(String name) {
        for (int i = 0 ; i < sheets.size() ; ++i) {
            XSSFSheet sheet = sheets.get(i);
            if (name.equalsIgnoreCase(sheet.getSheetName())) {
                return i;
            }
        }
        return -1;
    }

    /**
     * Returns the index of the given sheet
     *
     * @param sheet the sheet to look up
     * @return index of the sheet (0 based). <tt>-1</tt> if not found
     */
    public int getSheetIndex(Sheet sheet) {
        int idx = 0;
        for(XSSFSheet sh : sheets){
            if(sh == sheet) return idx;
            idx++;
        }
        return -1;
    }

    /**
     * Get the sheet name
     *
     * @param sheetIx Number
     * @return Sheet name
     */
    public String getSheetName(int sheetIx) {
        validateSheetIndex(sheetIx);
        return sheets.get(sheetIx).getSheetName();
    }

    /**
     * Allows foreach loops:
     * <pre><code>
     * XSSFWorkbook wb = new XSSFWorkbook(package);
     * for(XSSFSheet sheet : wb){
     *
     * }
     * </code></pre>
     */
    public Iterator<XSSFSheet> iterator() {
        return sheets.iterator();
    }
    /**
     * Are we a normal workbook (.xlsx), or a
     *  macro enabled workbook (.xlsm)?
     */
    public boolean isMacroEnabled() {
        return getPackagePart().getContentType().equals(XSSFRelation.MACROS_WORKBOOK.getContentType());
    }

    public void removeName(int nameIndex) {
        namedRanges.remove(nameIndex);
    }

    public void removeName(String name) {
        for (int i = 0; i < namedRanges.size(); i++) {
            XSSFName nm = namedRanges.get(i);
            if(nm.getNameName().equalsIgnoreCase(name)) {
                removeName(i);
                return;
            }
        }
        throw new IllegalArgumentException("Named range was not found: " + name);
    }

    /**
     * Delete the printarea for the sheet specified
     *
     * @param sheetIndex 0-based sheet index (0 = First Sheet)
     */
    public void removePrintArea(int sheetIndex) {
        int cont = 0;
        for (XSSFName name : namedRanges) {
            if (name.getNameName().equals(XSSFName.BUILTIN_PRINT_AREA) && name.getSheetIndex() == sheetIndex) {
                namedRanges.remove(cont);
                break;
            }
            cont++;
        }
    }

    /**
     * Removes sheet at the given index.<p/>
     *
     * Care must be taken if the removed sheet is the currently active or only selected sheet in
     * the workbook. There are a few situations when Excel must have a selection and/or active
     * sheet. (For example when printing - see Bug 40414).<br/>
     *
     * This method makes sure that if the removed sheet was active, another sheet will become
     * active in its place.  Furthermore, if the removed sheet was the only selected sheet, another
     * sheet will become selected.  The newly active/selected sheet will have the same index, or
     * one less if the removed sheet was the last in the workbook.
     *
     * @param index of the sheet  (0-based)
     */
    public void removeSheetAt(int index) {
        validateSheetIndex(index);

        onSheetDelete(index);

        XSSFSheet sheet = getSheetAt(index);
        removeRelation(sheet);
        sheets.remove(index);
    }

    /**
     * Gracefully remove references to the sheet being deleted
     *
     * @param index the 0-based index of the sheet to delete
     */
    private void onSheetDelete(int index) {
        //delete the CTSheet reference from workbook.xml
        workbook.getSheets().removeSheet(index);

        //calculation chain is auxiliary, remove it as it may contain orphan references to deleted cells
        if(calcChain != null) {
            removeRelation(calcChain);
            calcChain = null;
        }

        //adjust indices of names ranges
        for (Iterator<XSSFName> it = namedRanges.iterator(); it.hasNext();) {
            XSSFName nm = it.next();
            CTDefinedName ct = nm.getCTName();
            if(!ct.isSetLocalSheetId()) continue;
            if (ct.getLocalSheetId() == index) {
                it.remove();
            } else if (ct.getLocalSheetId() > index){
                // Bump down by one, so still points at the same sheet
                ct.setLocalSheetId(ct.getLocalSheetId()-1);
            }
        }
    }

    /**
     * Retrieves the current policy on what to do when
     *  getting missing or blank cells from a row.
     * The default is to return blank and null cells.
     *  {@link MissingCellPolicy}
     */
    public MissingCellPolicy getMissingCellPolicy() {
        return _missingCellPolicy;
    }
    /**
     * Sets the policy on what to do when
     *  getting missing or blank cells from a row.
     * This will then apply to all calls to
     *  {@link Row#getCell(int)}}. See
     *  {@link MissingCellPolicy}
     */
    public void setMissingCellPolicy(MissingCellPolicy missingCellPolicy) {
        _missingCellPolicy = missingCellPolicy;
    }

    /**
     * Convenience method to set the active sheet.  The active sheet is is the sheet
     * which is currently displayed when the workbook is viewed in Excel.
     * 'Selected' sheet(s) is a distinct concept.
     */
    @SuppressWarnings("deprecation") //YK: getXYZArray() array accessors are deprecated in xmlbeans with JDK 1.5 support
    public void setActiveSheet(int index) {

        validateSheetIndex(index);

        for (CTBookView arrayBook : workbook.getBookViews().getWorkbookViewArray()) {
            arrayBook.setActiveTab(index);
        }
    }

    /**
     * Validate sheet index
     *
     * @param index the index to validate
     * @throws IllegalArgumentException if the index is out of range (index
     *            &lt; 0 || index &gt;= getNumberOfSheets()).
    */
    private void validateSheetIndex(int index) {
        int lastSheetIx = sheets.size() - 1;
        if (index < 0 || index > lastSheetIx) {
            throw new IllegalArgumentException("Sheet index ("
                    + index +") is out of range (0.." +    lastSheetIx + ")");
        }
    }

    /**
     * Gets the first tab that is displayed in the list of tabs in excel.
     *
     * @return integer that contains the index to the active sheet in this book view.
     */
    public int getFirstVisibleTab() {
        CTBookViews bookViews = workbook.getBookViews();
        CTBookView bookView = bookViews.getWorkbookViewArray(0);
        return (short) bookView.getActiveTab();
    }

    /**
     * Sets the first tab that is displayed in the list of tabs in excel.
     *
     * @param index integer that contains the index to the active sheet in this book view.
     */
    public void setFirstVisibleTab(int index) {
        CTBookViews bookViews = workbook.getBookViews();
        CTBookView bookView= bookViews.getWorkbookViewArray(0);
        bookView.setActiveTab(index);
    }

    /**
     * Sets the printarea for the sheet provided
     * <p>
     * i.e. Reference = $A$1:$B$2
     * @param sheetIndex Zero-based sheet index (0 Represents the first sheet to keep consistent with java)
     * @param reference Valid name Reference for the Print Area
     */
    public void setPrintArea(int sheetIndex, String reference) {
        XSSFName name = getBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
        if (name == null) {
            name = createBuiltInName(XSSFName.BUILTIN_PRINT_AREA, sheetIndex);
        }
        //short externSheetIndex = getWorkbook().checkExternSheet(sheetIndex);
        //name.setExternSheetNumber(externSheetIndex);
        String[] parts = COMMA_PATTERN.split(reference);
        StringBuffer sb = new StringBuffer(32);
        for (int i = 0; i < parts.length; i++) {
            if(i>0) {
                sb.append(",");
            }
            SheetNameFormatter.appendFormat(sb, getSheetName(sheetIndex));
            sb.append("!");
            sb.append(parts[i]);
        }
        name.setRefersToFormula(sb.toString());
    }

    /**
     * For the Convenience of Java Programmers maintaining pointers.
     * @see #setPrintArea(int, String)
     * @param sheetIndex Zero-based sheet index (0 = First Sheet)
     * @param startColumn Column to begin printarea
     * @param endColumn Column to end the printarea
     * @param startRow Row to begin the printarea
     * @param endRow Row to end the printarea
     */
    public void setPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow) {
        String reference=getReferencePrintArea(getSheetName(sheetIndex), startColumn, endColumn, startRow, endRow);
        setPrintArea(sheetIndex, reference);
    }


    /**
     * Sets the repeating rows and columns for a sheet.
     * <p/>
     * To set just repeating columns:
     * <pre>
     *  workbook.setRepeatingRowsAndColumns(0,0,1,-1,-1);
     * </pre>
     * To set just repeating rows:
     * <pre>
     *  workbook.setRepeatingRowsAndColumns(0,-1,-1,0,4);
     * </pre>
     * To remove all repeating rows and columns for a sheet.
     * <pre>
     *  workbook.setRepeatingRowsAndColumns(0,-1,-1,-1,-1);
     * </pre>
     *
     * @param sheetIndex  0 based index to sheet.
     * @param startColumn 0 based start of repeating columns.
     * @param endColumn   0 based end of repeating columns.
     * @param startRow    0 based start of repeating rows.
     * @param endRow      0 based end of repeating rows.
     */
    public void setRepeatingRowsAndColumns(int sheetIndex,
                                           int startColumn, int endColumn,
                                           int startRow, int endRow) {
        //    Check arguments
        if ((startColumn == -1 && endColumn != -1) || startColumn < -1 || endColumn < -1 || startColumn > endColumn)
            throw new IllegalArgumentException("Invalid column range specification");
        if ((startRow == -1 && endRow != -1) || startRow < -1 || endRow < -1 || startRow > endRow)
            throw new IllegalArgumentException("Invalid row range specification");

        XSSFSheet sheet = getSheetAt(sheetIndex);
        boolean removingRange = startColumn == -1 && endColumn == -1 && startRow == -1 && endRow == -1;

        XSSFName name = getBuiltInName(XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
        if (removingRange) {
            if(name != null)namedRanges.remove(name);
            return;
        }
        if (name == null) {
            name = createBuiltInName(XSSFName.BUILTIN_PRINT_TITLE, sheetIndex);
        }

        String reference = getReferenceBuiltInRecord(name.getSheetName(), startColumn, endColumn, startRow, endRow);
        name.setRefersToFormula(reference);

        XSSFPrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setValidSettings(false);
    }

    private static String getReferenceBuiltInRecord(String sheetName, int startC, int endC, int startR, int endR) {
        //windows excel example for built-in title: 'second sheet'!$E:$F,'second sheet'!$2:$3
        CellReference colRef = new CellReference(sheetName, 0, startC, true, true);
        CellReference colRef2 = new CellReference(sheetName, 0, endC, true, true);

        String escapedName = SheetNameFormatter.format(sheetName);

        String c;
        if(startC == -1 && endC == -1) c= "";
        else c = escapedName + "!$" + colRef.getCellRefParts()[2] + ":$" + colRef2.getCellRefParts()[2];

        CellReference rowRef = new CellReference(sheetName, startR, 0, true, true);
        CellReference rowRef2 = new CellReference(sheetName, endR, 0, true, true);

        String r = "";
        if(startR == -1 && endR == -1) r = "";
        else {
            if (!rowRef.getCellRefParts()[1].equals("0") && !rowRef2.getCellRefParts()[1].equals("0")) {
                r = escapedName + "!$" + rowRef.getCellRefParts()[1] + ":$" + rowRef2.getCellRefParts()[1];
            }
        }

        StringBuffer rng = new StringBuffer();
        rng.append(c);
        if(rng.length() > 0 && r.length() > 0) rng.append(',');
        rng.append(r);
        return rng.toString();
    }

    private static String getReferencePrintArea(String sheetName, int startC, int endC, int startR, int endR) {
        //windows excel example: Sheet1!$C$3:$E$4
        CellReference colRef = new CellReference(sheetName, startR, startC, true, true);
        CellReference colRef2 = new CellReference(sheetName, endR, endC, true, true);

        return "$" + colRef.getCellRefParts()[2] + "$" + colRef.getCellRefParts()[1] + ":$" + colRef2.getCellRefParts()[2] + "$" + colRef2.getCellRefParts()[1];
    }

    public XSSFName getBuiltInName(String builtInCode, int sheetNumber) {
        for (XSSFName name : namedRanges) {
            if (name.getNameName().equalsIgnoreCase(builtInCode) && name.getSheetIndex() == sheetNumber) {
                return name;
            }
        }
        return null;
    }

    /**
     * Generates a NameRecord to represent a built-in region
     *
     * @return a new NameRecord
     * @throws IllegalArgumentException if sheetNumber is invalid
     * @throws POIXMLException if such a name already exists in the workbook
     */
    XSSFName createBuiltInName(String builtInName, int sheetNumber) {
        validateSheetIndex(sheetNumber);

        CTDefinedNames names = workbook.getDefinedNames() == null ? workbook.addNewDefinedNames() : workbook.getDefinedNames();
        CTDefinedName nameRecord = names.addNewDefinedName();
        nameRecord.setName(builtInName);
        nameRecord.setLocalSheetId(sheetNumber);

        XSSFName name = new XSSFName(nameRecord, this);
        for (XSSFName nr : namedRanges) {
            if (nr.equals(name))
                throw new POIXMLException("Builtin (" + builtInName
                        + ") already exists for sheet (" + sheetNumber + ")");
        }

        namedRanges.add(name);
        return name;
    }

    /**
     * We only set one sheet as selected for compatibility with HSSF.
     */
    public void setSelectedTab(int index) {
        for (int i = 0 ; i < sheets.size() ; ++i) {
            XSSFSheet sheet = sheets.get(i);
            sheet.setSelected(i == index);
        }
    }

    /**
     * Set the sheet name.
     * Will throw IllegalArgumentException if the name is greater than 31 chars
     * or contains /\?*[]
     *
     * @param sheetIndex number (0 based)
     */
    public void setSheetName(int sheetIndex, String name) {
        validateSheetIndex(sheetIndex);
        WorkbookUtil.validateSheetName(name);
        if (containsSheet(name, sheetIndex ))
            throw new IllegalArgumentException( "The workbook already contains a sheet of this name" );

/*        XSSFFormulaUtils utils = new XSSFFormulaUtils(this);
        utils.updateSheetName(sheetIndex, name);

        workbook.getSheets().getSheetArray(sheetIndex).setName(name);
*/        
        //20110106, henrichen@zkoss.org: handle the externsheet reference
        final Sheet wsheet = getSheetAt(sheetIndex);
        if (wsheet != null) {
	        final String oldname = wsheet.getSheetName();
	        for(String[] names : _externalSheetRefs) {
	        	final String sheetname1 = names[1];
	        	final String sheetname2 = names[2];
	        	if (oldname.equals(sheetname1)) {
	        		names[1] = name;
	        	}
	        	if (oldname.equals(sheetname2)) {
	        		names[2] = name;
	        	}
	        }
	        //20110112, henrichen@zkoss.org: adjust sheet name of the named range
			final String o = SheetNameFormatter.format(oldname);
			final String n = SheetNameFormatter.format(name);
	        for (XSSFName nm : namedRanges) {
	            final CTDefinedName ct = nm.getCTName();
	            if(ct.isSetLocalSheetId()) {
		            if (ct.getLocalSheetId() == sheetIndex) {
		            	final String ref = ct.getStringValue();
		            	final String newref = ref.replaceAll(o+"!", n+"!");
		            	ct.setStringValue(newref);
		            }
	            }
	        }
        }
        workbook.getSheets().getSheetArray(sheetIndex).setName(name);
    }

    /**
     * sets the order of appearance for a given sheet.
     *
     * @param sheetname the name of the sheet to reorder
     * @param pos the position that we want to insert the sheet into (0 based)
     */
    public void setSheetOrder(String sheetname, int pos) {
        int idx = getSheetIndex(sheetname);
        sheets.add(pos, sheets.remove(idx));
        // Reorder CTSheets
        CTSheets ct = workbook.getSheets();
        XmlObject cts = ct.getSheetArray(idx).copy();
        workbook.getSheets().removeSheet(idx);
        CTSheet newcts = ct.insertNewSheet(pos);
        newcts.set(cts);

        //notify sheets
        for(int i=0; i < sheets.size(); i++) {
            sheets.get(i).sheet = ct.getSheetArray(i);
        }
    }

    /**
     * marshal named ranges from the {@link #namedRanges} collection to the underlying CTWorkbook bean
     */
    private void saveNamedRanges(){
        // Named ranges
        if(namedRanges.size() > 0) {
            CTDefinedNames names = CTDefinedNames.Factory.newInstance();
            CTDefinedName[] nr = new CTDefinedName[namedRanges.size()];
            int i = 0;
            for(XSSFName name : namedRanges) {
                nr[i] = name.getCTName();
                i++;
            }
            names.setDefinedNameArray(nr);
            workbook.setDefinedNames(names); 
            
            //bug#ZSS-36: Exception when exporting excel twice.
            //20110818, henrichen: names.setDefinedNameArray() will instantiate new CTDefinedNames 
            // and those in namedRanged is orphaned. Have to sync back CTDefinedName into namedRanges
            syncNamedRange();
        } else {
            if(workbook.isSetDefinedNames()) {
                workbook.unsetDefinedNames();
            }
        }
    }

    private void saveCalculationChain(){
        if(calcChain != null){
            int count = calcChain.getCTCalcChain().sizeOfCArray();
            if(count == 0){
                removeRelation(calcChain);
                calcChain = null;
            }
        }
    }

    @Override
    protected void commit() throws IOException {
        saveNamedRanges();
        saveCalculationChain();

        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        xmlOptions.setSaveSyntheticDocumentElement(new QName(CTWorkbook.type.getName().getNamespaceURI(), "workbook"));
        Map<String, String> map = new HashMap<String, String>();
        map.put(STRelationshipId.type.getName().getNamespaceURI(), "r");
        xmlOptions.setSaveSuggestedPrefixes(map);

        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        workbook.save(out, xmlOptions);
        out.close();
    }

    /**
     * Returns SharedStringsTable - tha cache of string for this workbook
     *
     * @return the shared string table
     */
    @Internal
    public SharedStringsTable getSharedStringSource() {
        return this.sharedStringSource;
    }

    /**
     * Return a object representing a collection of shared objects used for styling content,
     * e.g. fonts, cell styles, colors, etc.
     */
    public StylesTable getStylesSource() {
        return this.stylesSource;
    }

    /**
     * Returns the Theme of current workbook.
     */
    public ThemesTable getTheme() {
        return theme;
    }

    /**
     * Returns an object that handles instantiating concrete
     *  classes of the various instances for XSSF.
     */
    public XSSFCreationHelper getCreationHelper() {
        if(_creationHelper == null) _creationHelper = new XSSFCreationHelper(this);
        return _creationHelper;
    }

    /**
     * Determines whether a workbook contains the provided sheet name.
     * For the purpose of comparison, long names are truncated to 31 chars.
     *
     * @param name the name to test (case insensitive match)
     * @param excludeSheetIdx the sheet to exclude from the check or -1 to include all sheets in the check.
     * @return true if the sheet contains the name, false otherwise.
     */
    @SuppressWarnings("deprecation") //  getXYZArray() array accessors are deprecated
    private boolean containsSheet(String name, int excludeSheetIdx) {
        CTSheet[] ctSheetArray = workbook.getSheets().getSheetArray();

        if (name.length() > MAX_SENSITIVE_SHEET_NAME_LEN) {
            name = name.substring(0, MAX_SENSITIVE_SHEET_NAME_LEN);
        }

        for (int i = 0; i < ctSheetArray.length; i++) {
            String ctName = ctSheetArray[i].getName();
            if (ctName.length() > MAX_SENSITIVE_SHEET_NAME_LEN) {
                ctName = ctName.substring(0, MAX_SENSITIVE_SHEET_NAME_LEN);
            }

            if (excludeSheetIdx != i && name.equalsIgnoreCase(ctName))
                return true;
        }
        return false;
    }

    /**
     * Gets a boolean value that indicates whether the date systems used in the workbook starts in 1904.
     * <p>
     * The default value is false, meaning that the workbook uses the 1900 date system,
     * where 1/1/1900 is the first day in the system..
     * </p>
     * @return true if the date systems used in the workbook starts in 1904
     */
    protected boolean isDate1904(){
        CTWorkbookPr workbookPr = workbook.getWorkbookPr();
        return workbookPr != null && workbookPr.getDate1904();
    }

    /**
     * Get the document's embedded files.
     */
    public List<PackagePart> getAllEmbedds() throws OpenXML4JException {
        List<PackagePart> embedds = new LinkedList<PackagePart>();

        for(XSSFSheet sheet : sheets){
            // Get the embeddings for the workbook
            for(PackageRelationship rel : sheet.getPackagePart().getRelationshipsByType(XSSFRelation.OLEEMBEDDINGS.getRelation()))
                embedds.add(getTargetPart(rel));

            for(PackageRelationship rel : sheet.getPackagePart().getRelationshipsByType(XSSFRelation.PACKEMBEDDINGS.getRelation()))
                embedds.add(getTargetPart(rel));

        }
        return embedds;
    }

    public boolean isHidden() {
        throw new RuntimeException("Not implemented yet");
    }

    public void setHidden(boolean hiddenFlag) {
        throw new RuntimeException("Not implemented yet");
    }

    /**
     * Check whether a sheet is hidden.
     * <p>
     * Note that a sheet could instead be set to be very hidden, which is different
     *  ({@link #isSheetVeryHidden(int)})
     * </p>
     * @param sheetIx Number
     * @return <code>true</code> if sheet is hidden
     */
    public boolean isSheetHidden(int sheetIx) {
        validateSheetIndex(sheetIx);
        CTSheet ctSheet = sheets.get(sheetIx).sheet;
        return ctSheet.getState() == STSheetState.HIDDEN;
    }

    /**
     * Check whether a sheet is very hidden.
     * <p>
     * This is different from the normal hidden status
     *  ({@link #isSheetHidden(int)})
     * </p>
     * @param sheetIx sheet index to check
     * @return <code>true</code> if sheet is very hidden
     */
    public boolean isSheetVeryHidden(int sheetIx) {
        validateSheetIndex(sheetIx);
        CTSheet ctSheet = sheets.get(sheetIx).sheet;
        return ctSheet.getState() == STSheetState.VERY_HIDDEN;
    }

    /**
     * Sets the visible state of this sheet.
     * <p>
     *   Calling <code>setSheetHidden(sheetIndex, true)</code> is equivalent to
     *   <code>setSheetHidden(sheetIndex, Workbook.SHEET_STATE_HIDDEN)</code>.
     * <br/>
     *   Calling <code>setSheetHidden(sheetIndex, false)</code> is equivalent to
     *   <code>setSheetHidden(sheetIndex, Workbook.SHEET_STATE_VISIBLE)</code>.
     * </p>
     *
     * @param sheetIx   the 0-based index of the sheet
     * @param hidden whether this sheet is hidden
     * @see #setSheetHidden(int, int)
     */
    public void setSheetHidden(int sheetIx, boolean hidden) {
        setSheetHidden(sheetIx, hidden ? SHEET_STATE_HIDDEN : SHEET_STATE_VISIBLE);
    }

    /**
     * Hide or unhide a sheet.
     *
     * <ul>
     *  <li>0 - visible. </li>
     *  <li>1 - hidden. </li>
     *  <li>2 - very hidden.</li>
     * </ul>
     * @param sheetIx the sheet index (0-based)
     * @param state one of the following <code>Workbook</code> constants:
     *        <code>Workbook.SHEET_STATE_VISIBLE</code>,
     *        <code>Workbook.SHEET_STATE_HIDDEN</code>, or
     *        <code>Workbook.SHEET_STATE_VERY_HIDDEN</code>.
     * @throws IllegalArgumentException if the supplied sheet index or state is invalid
     */
    public void setSheetHidden(int sheetIx, int state) {
        validateSheetIndex(sheetIx);
        WorkbookUtil.validateSheetState(state);
        CTSheet ctSheet = sheets.get(sheetIx).sheet;
        ctSheet.setState(STSheetState.Enum.forInt(state + 1));
    }

    /**
     * Fired when a formula is deleted from this workbook,
     * for example when calling cell.setCellFormula(null)
     *
     * @see XSSFCell#setCellFormula(String)
     */
    protected void onDeleteFormula(XSSFCell cell){
        if(calcChain != null) {
            int sheetId = (int)cell.getSheet().sheet.getSheetId();
            calcChain.removeItem(sheetId, cell.getReference());
        }
    }

    /**
     * Return the CalculationChain object for this workbook
     * <p>
     *   The calculation chain object specifies the order in which the cells in a workbook were last calculated
     * </p>
     *
     * @return the <code>CalculationChain</code> object or <code>null</code> if not defined
     */
    @Internal
    public CalculationChain getCalculationChain(){
        return calcChain;
    }

    /**
     *
     * @return a collection of custom XML mappings defined in this workbook
     */
    public Collection<XSSFMap> getCustomXMLMappings(){
        return mapInfo == null ? new ArrayList<XSSFMap>() : mapInfo.getAllXSSFMaps();
    }

    /**
     *
     * @return the helper class used to query the custom XML mapping defined in this workbook
     */
    @Internal
    public MapInfo getMapInfo(){
    	return mapInfo;
    }


	/**
	 * Specifies a boolean value that indicates whether structure of workbook is locked. <br/>
	 * A value true indicates the structure of the workbook is locked. Worksheets in the workbook can't be moved,
	 * deleted, hidden, unhidden, or renamed, and new worksheets can't be inserted.<br/>
	 * A value of false indicates the structure of the workbook is not locked.<br/>
	 * 
	 * @return true if structure of workbook is locked
	 */
	public boolean isStructureLocked() {
		return workbookProtectionPresent() && workbook.getWorkbookProtection().getLockStructure();
	}

	/**
	 * Specifies a boolean value that indicates whether the windows that comprise the workbook are locked. <br/>
	 * A value of true indicates the workbook windows are locked. Windows are the same size and position each time the
	 * workbook is opened.<br/>
	 * A value of false indicates the workbook windows are not locked.
	 * 
	 * @return true if windows that comprise the workbook are locked
	 */
	public boolean isWindowsLocked() {
		return workbookProtectionPresent() && workbook.getWorkbookProtection().getLockWindows();
	}

	/**
	 * Specifies a boolean value that indicates whether the workbook is locked for revisions.
	 * 
	 * @return true if the workbook is locked for revisions.
	 */
	public boolean isRevisionLocked() {
		return workbookProtectionPresent() && workbook.getWorkbookProtection().getLockRevision();
	}
	
	/**
	 * Locks the structure of workbook.
	 */
	public void lockStructure() {
		createProtectionFieldIfNotPresent();
		workbook.getWorkbookProtection().setLockStructure(true);
	}
	
	/**
	 * Unlocks the structure of workbook.
	 */
	public void unLockStructure() {
		createProtectionFieldIfNotPresent();
		workbook.getWorkbookProtection().setLockStructure(false);
	}

	/**
	 * Locks the windows that comprise the workbook. 
	 */
	public void lockWindows() {
		createProtectionFieldIfNotPresent();
		workbook.getWorkbookProtection().setLockWindows(true);
	}
	
	/**
	 * Unlocks the windows that comprise the workbook. 
	 */
	public void unLockWindows() {
		createProtectionFieldIfNotPresent();
		workbook.getWorkbookProtection().setLockWindows(false);
	}
	
	/**
	 * Locks the workbook for revisions.
	 */
	public void lockRevision() {
		createProtectionFieldIfNotPresent();
		workbook.getWorkbookProtection().setLockRevision(true);
	}

	/**
	 * Unlocks the workbook for revisions.
	 */
	public void unLockRevision() {
		createProtectionFieldIfNotPresent();
		workbook.getWorkbookProtection().setLockRevision(false);
	}
	
	private boolean workbookProtectionPresent() {
		return workbook.getWorkbookProtection() != null;
	}

	private void createProtectionFieldIfNotPresent() {
		if (workbook.getWorkbookProtection() == null){
			workbook.setWorkbookProtection(CTWorkbookProtection.Factory.newInstance());
		}
	}

    /**
     *
     * Returns the locator of user-defined functions.
     * <p>
     * The default instance extends the built-in functions with the Excel Analysis Tool Pack.
     * To set / evaluate custom functions you need to register them as follows:
     *
     *
     *
     * </p>
     * @return wrapped instance of UDFFinder that allows seeking functions both by index and name
     */
    /*package*/ UDFFinder getUDFFinder() {
        return _udfFinder;
    }

    /**
     * Register a new toolpack in this workbook.
     *
     * @param toopack the toolpack to register
     */
    public void addToolPack(UDFFinder toopack){
        _udfFinder.add(toopack);
    }



	/*package*/ String getBookNameFromExternalLinkIndex(String externalLinkIndex) {
		return linkIndexToBookName.get(externalLinkIndex);
	}
	
	//20110818, henrichen: sync CTDefinedName and namedRanges
	private void syncNamedRange() {
        namedRanges = new ArrayList<XSSFName>();
        if(workbook.isSetDefinedNames()) {
            for(CTDefinedName ctName : workbook.getDefinedNames().getDefinedNameArray()) {
                namedRanges.add(new XSSFName(ctName, this));
            }
        }
	}
}
