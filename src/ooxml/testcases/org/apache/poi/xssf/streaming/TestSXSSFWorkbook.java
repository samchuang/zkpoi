/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package org.apache.poi.xssf.streaming;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.SXSSFITestDataProvider;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;

public final class TestSXSSFWorkbook extends BaseTestWorkbook {

	public TestSXSSFWorkbook() {
		super(SXSSFITestDataProvider.instance);
	}

    /**
     * cloning of sheets is not supported in SXSSF
     */
    @Override
    public void testCloneSheet() {
        try {
            super.testCloneSheet();
            fail("expected exception");
        } catch (RuntimeException e){
            assertEquals("NotImplemented", e.getMessage());
        }
    }

    /**
     * this test involves evaluation of formulas which isn't supported for SXSSF
     */
    @Override
    public void testSetSheetName() {
        try {
            super.testSetSheetName();
            fail("expected exception");
        } catch (Exception e){
            assertEquals(
                    "Unexpected type of cell: class org.apache.poi.xssf.streaming.SXSSFCell. " +
                    "Only XSSFCells can be evaluated.", e.getMessage());
        }
    }
    
    public void testExistingWorkbook() {
    	XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
    	xssfWorkbook.createSheet("S1");
    	SXSSFWorkbook wb = new SXSSFWorkbook(xssfWorkbook);
    	xssfWorkbook = (XSSFWorkbook) SXSSFITestDataProvider.instance.writeOutAndReadBack(wb);
    	wb = new SXSSFWorkbook(xssfWorkbook);
    	assertEquals(1, wb.getNumberOfSheets());
    	Sheet sheet  = wb.getSheetAt(0);
    	assertNotNull(sheet);
    	assertEquals("S1", sheet.getSheetName());
    }
    
    public void testAddToExistingWorkbook() {
    	XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
    	xssfWorkbook.createSheet("S1");
    	Sheet sheet = xssfWorkbook.createSheet("S2");
    	Row row = sheet.createRow(1);
    	Cell cell = row.createCell(1);
    	cell.setCellValue("value 2_1_1");
    	SXSSFWorkbook wb = new SXSSFWorkbook(xssfWorkbook);
    	xssfWorkbook = (XSSFWorkbook) SXSSFITestDataProvider.instance.writeOutAndReadBack(wb);
    	wb = new SXSSFWorkbook(xssfWorkbook);
    	
    	// Add a row to the existing empty sheet
    	Sheet sheet1 = wb.getSheetAt(0);
    	Row row1_1 = sheet1.createRow(1);
    	Cell cell1_1_1 = row1_1.createCell(1);
    	cell1_1_1.setCellValue("value 1_1_1");
    	
    	// Add a row to the existing non-empty sheet
    	Sheet sheet2 = wb.getSheetAt(1);
    	Row row2_2 = sheet2.createRow(2);
    	Cell cell2_2_1 = row2_2.createCell(1);
    	cell2_2_1.setCellValue("value 2_2_1");
    	
    	// Add a sheet with one row
    	Sheet sheet3 = wb.createSheet("S3");
    	Row row3_1 = sheet3.createRow(1);
    	Cell cell3_1_1 = row3_1.createCell(1);
    	cell3_1_1.setCellValue("value 3_1_1");
    	
    	xssfWorkbook = (XSSFWorkbook) SXSSFITestDataProvider.instance.writeOutAndReadBack(wb);
    	assertEquals(3, xssfWorkbook.getNumberOfSheets());
    	// Verify sheet 1
    	sheet1 = xssfWorkbook.getSheetAt(0);
    	assertEquals("S1", sheet1.getSheetName());
    	assertEquals(1, sheet1.getPhysicalNumberOfRows());
    	row1_1 = sheet1.getRow(1);
    	assertNotNull(row1_1);
    	cell1_1_1 = row1_1.getCell(1);
    	assertNotNull(cell1_1_1);
    	assertEquals("value 1_1_1", cell1_1_1.getStringCellValue());
    	// Verify sheet 2
    	sheet2 = xssfWorkbook.getSheetAt(1);
    	assertEquals("S2", sheet2.getSheetName());
    	assertEquals(2, sheet2.getPhysicalNumberOfRows());
    	Row row2_1 = sheet2.getRow(1);
    	assertNotNull(row2_1);
    	Cell cell2_1_1 = row2_1.getCell(1);
    	assertNotNull(cell2_1_1);
    	assertEquals("value 2_1_1", cell2_1_1.getStringCellValue());
    	row2_2 = sheet2.getRow(2);
    	assertNotNull(row2_2);
    	cell2_2_1 = row2_2.getCell(1);
    	assertNotNull(cell2_2_1);
    	assertEquals("value 2_2_1", cell2_2_1.getStringCellValue());
    	// Verify sheet 3
    	sheet3 = xssfWorkbook.getSheetAt(2);
    	assertEquals("S3", sheet3.getSheetName());
    	assertEquals(1, sheet3.getPhysicalNumberOfRows());
    	row3_1 = sheet3.getRow(1);
    	assertNotNull(row3_1);
    	cell3_1_1 = row3_1.getCell(1);
    	assertNotNull(cell3_1_1);
    	assertEquals("value 3_1_1", cell3_1_1.getStringCellValue());
    }

    public void testSheetdataWriter(){
        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFSheet sh = (SXSSFSheet)wb.createSheet();
        SheetDataWriter wr = sh.getSheetDataWriter();
        assertTrue(wr.getClass() == SheetDataWriter.class);
        File tmp = wr.getTempFile();
        assertTrue(tmp.getName().startsWith("poi-sxssf-sheet"));
        assertTrue(tmp.getName().endsWith(".xml"));

        wb = new SXSSFWorkbook();
        wb.setCompressTempFiles(true);
        sh = (SXSSFSheet)wb.createSheet();
        wr = sh.getSheetDataWriter();
        assertTrue(wr.getClass() == GZIPSheetDataWriter.class);
        tmp = wr.getTempFile();
        assertTrue(tmp.getName().startsWith("poi-sxssf-sheet-xml"));
        assertTrue(tmp.getName().endsWith(".gz"));
    }
    
    public void testGZipSheetdataWriter(){
        Workbook wb = new SXSSFWorkbook();
        ((SXSSFWorkbook)wb).setCompressTempFiles(true);
        int rowNum = 1000;
        int sheetNum = 5;
        for(int i = 0; i < sheetNum; i++){
            Sheet sh = wb.createSheet("sheet" + i);
            for(int j = 0; j < rowNum; j++){
                Row row = sh.createRow(j);
                Cell cell1 = row.createCell(0);
                cell1.setCellValue(new CellReference(cell1).formatAsString());
                
                Cell cell2 = row.createCell(1);
                cell2.setCellValue(i);

                Cell cell3 = row.createCell(2);
                cell3.setCellValue(j);
            }
        }

        wb = SXSSFITestDataProvider.instance.writeOutAndReadBack(wb);
        for(int i = 0; i < sheetNum; i++){
            Sheet sh = wb.getSheetAt(i);
            assertEquals("sheet" + i, sh.getSheetName());
            for(int j = 0; j < rowNum; j++){
                Row row = sh.getRow(j);
                assertNotNull("row[" + j + "]", row);
                Cell cell1 = row.getCell(0);
                assertEquals(new CellReference(cell1).formatAsString(), cell1.getStringCellValue());

                Cell cell2 = row.getCell(1);
                assertEquals(i, (int)cell2.getNumericCellValue());

                Cell cell3 = row.getCell(2);
                assertEquals(j, (int)cell3.getNumericCellValue());
            }
        }

        
    }

}
