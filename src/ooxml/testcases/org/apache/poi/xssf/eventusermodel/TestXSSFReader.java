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

package org.apache.poi.xssf.eventusermodel;

import java.io.InputStream;
import java.util.Iterator;

import junit.framework.TestCase;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.XSSFTestDataSamples;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.POIDataSamples;

/**
 * Tests for {@link XSSFReader}
 */
public final class TestXSSFReader extends TestCase {
    private static POIDataSamples _ssTests = POIDataSamples.getSpreadSheetInstance();

    public void testGetBits() throws Exception {
		OPCPackage pkg = OPCPackage.open(_ssTests.openResourceAsStream("SampleSS.xlsx"));

		XSSFReader r = new XSSFReader(pkg);

		assertNotNull(r.getWorkbookData());
		assertNotNull(r.getSharedStringsData());
		assertNotNull(r.getStylesData());

		assertNotNull(r.getSharedStringsTable());
		assertNotNull(r.getStylesTable());
	}

	public void testStyles() throws Exception {
		OPCPackage pkg = OPCPackage.open(_ssTests.openResourceAsStream("SampleSS.xlsx"));

		XSSFReader r = new XSSFReader(pkg);

		assertEquals(3, r.getStylesTable().getFonts().size());
		assertEquals(0, r.getStylesTable()._getNumberFormatSize());
		
		// The Styles Table should have the themes associated with it too
		assertNotNull(r.getStylesTable().getTheme());
		
		// Check we get valid data for the two
		assertNotNull(r.getStylesData());
      assertNotNull(r.getThemesData());
	}

	public void testStrings() throws Exception {
        OPCPackage pkg = OPCPackage.open(_ssTests.openResourceAsStream("SampleSS.xlsx"));

		XSSFReader r = new XSSFReader(pkg);

		assertEquals(11, r.getSharedStringsTable().getItems().size());
		assertEquals("Test spreadsheet", new XSSFRichTextString(r.getSharedStringsTable().getEntryAt(0)).toString());
	}

	public void testSheets() throws Exception {
        OPCPackage pkg = OPCPackage.open(_ssTests.openResourceAsStream("SampleSS.xlsx"));

		XSSFReader r = new XSSFReader(pkg);
		byte[] data = new byte[4096];

		// By r:id
		assertNotNull(r.getSheet("rId2"));
		int read = IOUtils.readFully(r.getSheet("rId2"), data);
		assertEquals(974, read);

		// All
		Iterator<InputStream> it = r.getSheetsData();

		int count = 0;
		while(it.hasNext()) {
			count++;
			InputStream inp = it.next();
			assertNotNull(inp);
			read = IOUtils.readFully(inp, data);
			inp.close();

			assertTrue(read > 400);
			assertTrue(read < 1500);
		}
		assertEquals(3, count);
	}

	/**
	 * Check that the sheet iterator returns sheets in the logical order
	 * (as they are defined in the workbook.xml)
	 */
	public void testOrderOfSheets() throws Exception {
        OPCPackage pkg = OPCPackage.open(_ssTests.openResourceAsStream("reordered_sheets.xlsx"));

		XSSFReader r = new XSSFReader(pkg);

		String[] sheetNames = {"Sheet4", "Sheet2", "Sheet3", "Sheet1"};
		XSSFReader.SheetIterator it = (XSSFReader.SheetIterator)r.getSheetsData();

		int count = 0;
		while(it.hasNext()) {
			InputStream inp = it.next();
			assertNotNull(inp);
			inp.close();

			assertEquals(sheetNames[count], it.getSheetName());
			count++;
		}
		assertEquals(4, count);
	}
	
	public void testComments() throws Exception {
      OPCPackage pkg =  XSSFTestDataSamples.openSamplePackage("comments.xlsx");
      XSSFReader r = new XSSFReader(pkg);
      XSSFReader.SheetIterator it = (XSSFReader.SheetIterator)r.getSheetsData();
      
      int count = 0;
      while(it.hasNext()) {
         count++;
         InputStream inp = it.next();
         inp.close();

         if(count == 1) {
            assertNotNull(it.getSheetComments());
            CommentsTable ct = it.getSheetComments();
            assertEquals(1, ct.getNumberOfAuthors());
            assertEquals(3, ct.getNumberOfComments());
         } else {
            assertNull(it.getSheetComments());
         }
      }
      assertEquals(3, count);
	}
   
   /**
    * Iterating over a workbook with chart sheets in it, using the
    *  XSSFReader method
    * @throws Exception
    */
   public void test50119() throws Exception {
      OPCPackage pkg =  XSSFTestDataSamples.openSamplePackage("WithChartSheet.xlsx");
      XSSFReader r = new XSSFReader(pkg);
      XSSFReader.SheetIterator it = (XSSFReader.SheetIterator)r.getSheetsData();
      
      while(it.hasNext())
      {
          InputStream stream = it.next();
          stream.close();
      }
   }
}
