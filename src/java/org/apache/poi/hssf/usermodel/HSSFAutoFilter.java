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

package org.zkoss.poi.hssf.usermodel;
import java.util.List;

import org.zkoss.poi.hssf.record.aggregates.AutoFilterInfoRecordAggregate;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.poi.ss.usermodel.FilterColumn;
import org.zkoss.poi.ss.util.CellRangeAddress;

/**
 * Represents autofiltering for the specified worksheet.
 *
 * @author Peterkuo
 */
public final class HSSFAutoFilter implements AutoFilter {
    private HSSFSheet _sheet;

    private AutoFilterInfoRecordAggregate _record;
    
    HSSFAutoFilter(HSSFSheet sheet, AutoFilterInfoRecordAggregate record){
        _sheet = sheet;
        _record = record;
    }

	public List<String> getValuesOfFilter(int column) {
		
		return _record.getValuesOfFilter(column);
	}

	@Override
	public CellRangeAddress getRangeAddress() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public List<FilterColumn> getFilterColumns() {
		// TODO Auto-generated method stub
		return null;
	}
	
	public FilterColumn getFilterColumn(int col) {
		return null;
	}
}