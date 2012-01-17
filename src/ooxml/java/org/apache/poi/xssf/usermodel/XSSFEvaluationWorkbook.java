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

import org.zkoss.poi.ss.formula.functions.FreeRefFunction;
import org.zkoss.poi.ss.formula.ptg.NamePtg;
import org.zkoss.poi.ss.formula.ptg.NameXPtg;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.ss.SpreadsheetVersion;
import org.zkoss.poi.ss.formula.EvaluationCell;
import org.zkoss.poi.ss.formula.EvaluationName;
import org.zkoss.poi.ss.formula.EvaluationSheet;
import org.zkoss.poi.ss.formula.EvaluationWorkbook;
import org.zkoss.poi.ss.formula.FormulaParser;
import org.zkoss.poi.ss.formula.FormulaParsingWorkbook;
import org.zkoss.poi.ss.formula.FormulaRenderingWorkbook;
import org.zkoss.poi.ss.formula.FormulaType;
import org.zkoss.poi.ss.formula.udf.UDFFinder;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.xssf.model.IndexedUDFFinder;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;

/**
 * Internal POI use only
 *
 * @author Josh Micich
 * @author Henri Chen (henrichen at zkoss dot org) - Sheet1:Sheet3!xxx 3d reference
 */
public final class XSSFEvaluationWorkbook implements FormulaRenderingWorkbook, EvaluationWorkbook, FormulaParsingWorkbook {

	private final XSSFWorkbook _uBook;

	public static XSSFEvaluationWorkbook create(XSSFWorkbook book) {
		if (book == null) {
			return null;
		}
		return new XSSFEvaluationWorkbook(book);
	}

	private XSSFEvaluationWorkbook(XSSFWorkbook book) {
		_uBook = book;
	}

	/**
	 * @return the sheet index of the sheet with the given external index.
	 */
	public int convertFromExternSheetIndex(int externSheetIndex) {
		final String[] names = _uBook.convertFromExternSheetIndex(externSheetIndex);
		if (names == null) return -1;
		return getSheetIndex(names[1]);
	}
	/**
	 * @return the sheet index of the 2nd sheet with the given external index.
	 */
	public int convertLastIndexFromExternSheetIndex(int externSheetIndex) {
		final String[] names = _uBook.convertFromExternSheetIndex(externSheetIndex);
		if (names == null) return -1;
		return getSheetIndex(names[2]);
	}
	
	public int getExternalSheetIndex(String sheetName) {
		final int j = sheetName.indexOf(':');
		final String sheetName1 = j < 0 ? sheetName : sheetName.substring(0, j);
		final String sheetName2 = j < 0 ? sheetName : sheetName.substring(j+1);
		return _uBook.getOrCreateExternalSheetIndex(null, sheetName1, sheetName2);
	}

	public EvaluationName getName(String name, int sheetIndex) {
		for (int i = 0; i < _uBook.getNumberOfNames(); i++) {
			XSSFName nm = _uBook.getNameAt(i);
			String nameText = nm.getNameName();
			if (name.equalsIgnoreCase(nameText) && nm.getSheetIndex() == sheetIndex) {
				return new Name(_uBook.getNameAt(i), i, this);
			}
		}
		return sheetIndex == -1 ? null : getName(name, -1);
	}

	public int getSheetIndex(EvaluationSheet evalSheet) {
		XSSFSheet sheet = ((XSSFEvaluationSheet)evalSheet).getXSSFSheet();
		return _uBook.getSheetIndex(sheet);
	}

	public String getSheetName(int sheetIndex) {
		return _uBook.getSheetName(sheetIndex);
	}
	
	public ExternalName getExternalName(int externSheetIndex, int externNameIndex) {
	   throw new RuntimeException("Not implemented yet");
	}

	public NameXPtg getNameXPtg(String name) {
        IndexedUDFFinder udfFinder = (IndexedUDFFinder)getUDFFinder();
        FreeRefFunction func = udfFinder.findFunction(name);
		if(func == null) return null;
        else return new NameXPtg(0, udfFinder.getFunctionIndex(name));
	}

    public String resolveNameXText(NameXPtg n) {
        int idx = n.getNameIndex();
        IndexedUDFFinder udfFinder = (IndexedUDFFinder)getUDFFinder();
        return udfFinder.getFunctionName(idx);
    }

	public EvaluationSheet getSheet(int sheetIndex) {
		return new XSSFEvaluationSheet(_uBook.getSheetAt(sheetIndex));
	}

	/** Return null if in the same workbook */
	public ExternalSheet getExternalSheet(int externSheetIndex) {
		String[] names = _uBook.convertFromExternSheetIndex(externSheetIndex);
		if (names != null && names[0] != null) {
			return new ExternalSheet(names[0], names[1], names[2]);
		}
		return null;
	}
	public int getExternalSheetIndex(String workbookName, String sheetName) {
		final int j = sheetName.indexOf(':');
		final String sheetName1 = j < 0 ? sheetName : sheetName.substring(0, j);
		final String sheetName2 = j < 0 ? sheetName : sheetName.substring(j+1);
		return _uBook.getOrCreateExternalSheetIndex(workbookName, sheetName1, sheetName2);
	}
	public int getSheetIndex(String sheetName) {
		return _uBook.getSheetIndex(sheetName);
	}

	public String getSheetNameByExternSheet(int externSheetIndex) {
		String[] names = _uBook.convertFromExternSheetIndex(externSheetIndex);
		//20120117, henrichen@zkoss.org: ZSS-82
		return names == null ? "" : names[1].equals(names[2]) ? names[1] : names[1]+':'+names[2];   
	}

	public String getNameText(NamePtg namePtg) {
		return _uBook.getNameAt(namePtg.getIndex()).getNameName();
	}
	public EvaluationName getName(NamePtg namePtg) {
		int ix = namePtg.getIndex();
		return new Name(_uBook.getNameAt(ix), ix, this);
	}
	public Ptg[] getFormulaTokens(EvaluationCell evalCell) {
		XSSFCell cell = ((XSSFEvaluationCell)evalCell).getXSSFCell();
		XSSFEvaluationWorkbook frBook = XSSFEvaluationWorkbook.create(_uBook);
		return FormulaParser.parse(cell.getCellFormula(), frBook, FormulaType.CELL, _uBook.getSheetIndex(cell.getSheet()));
	}

    public UDFFinder getUDFFinder(){
        return _uBook.getUDFFinder();
    }

	private static final class Name implements EvaluationName {

		private final XSSFName _nameRecord;
		private final int _index;
		private final FormulaParsingWorkbook _fpBook;

		public Name(XSSFName name, int index, FormulaParsingWorkbook fpBook) {
			_nameRecord = name;
			_index = index;
			_fpBook = fpBook;
		}

		public Ptg[] getNameDefinition() {

			return FormulaParser.parse(_nameRecord.getRefersToFormula(), _fpBook, FormulaType.NAMEDRANGE, _nameRecord.getSheetIndex());
		}

		public String getNameText() {
			return _nameRecord.getNameName();
		}

		public boolean hasFormula() {
			// TODO - no idea if this is right
			CTDefinedName ctn = _nameRecord.getCTName();
			String strVal = ctn.getStringValue();
			return !ctn.getFunction() && strVal != null && strVal.length() > 0;
		}

		public boolean isFunctionName() {
			return _nameRecord.isFunctionName();
		}

		public boolean isRange() {
			return hasFormula(); // TODO - is this right?
		}
		public NamePtg createPtg() {
			return new NamePtg(_index);
		}
	}

	public SpreadsheetVersion getSpreadsheetVersion(){
		return SpreadsheetVersion.EXCEL2007;
	}

	@Override
	public String getBookNameFromExternalLinkIndex(String externalLinkIndex) {
		return _uBook.getBookNameFromExternalLinkIndex(externalLinkIndex);
	}
	
	//20101112, henrichen@zkoss.org: handle parsing of user defined function
	@Override
	public EvaluationName getOrCreateName(String name, int sheetIndex) {
		for (int i = 0; i < _uBook.getNumberOfNames(); i++) {
			XSSFName nm = _uBook.getNameAt(i);
			String nameText = nm.getNameName();
			if (name.equalsIgnoreCase(nameText) && nm.getSheetIndex() == sheetIndex) {
				return new Name(_uBook.getNameAt(i), i, this);
			}
		}
		if (sheetIndex == -1) {
			XSSFName nm = _uBook.createName();
			nm.setNameName(name);
			return new Name(nm, _uBook.getNumberOfNames() - 1, this);
		}
		return getOrCreateName(name, -1); //recursive
	}
	
	//20111124, henrichen@zkoss.org: get formula tokens per the formula string
	public Ptg[] getFormulaTokens(int sheetIndex, String formula) {
		XSSFEvaluationWorkbook frBook = XSSFEvaluationWorkbook.create(_uBook);
		return FormulaParser.parse(formula, frBook, FormulaType.CELL, sheetIndex);
	}

	//20110117, henrichen@zkoss.org: get extern book index from book name
	//ZSS-81 Cannot input formula with proper external book name
	@Override
	public String getExternalLinkIndexFromBookName(String bookname) {
		return _uBook.getExternalLinkIndexFromBookName(bookname);
	}
}
