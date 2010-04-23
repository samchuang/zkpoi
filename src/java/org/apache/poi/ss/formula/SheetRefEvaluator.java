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

package org.apache.poi.ss.formula;

import org.apache.poi.hssf.record.formula.eval.ValueEval;
/**
 *
 *
 * @author Josh Micich
 * @author Henri Chen (henrichen at zkoss dot org) - Sheet1:Sheet3!xxx 3d reference
 */
final class SheetRefEvaluator {

	private final WorkbookEvaluator _bookEvaluator;
	private final EvaluationTracker _tracker;
	private final int _sheetIndex;
	private final int _lastSheetIndex;

	public SheetRefEvaluator(WorkbookEvaluator bookEvaluator, EvaluationTracker tracker, int sheetIndex, int lastSheetIndex) {
		if (sheetIndex < 0) {
			throw new IllegalArgumentException("Invalid sheetIndex: " + sheetIndex + ".");
		}
		if (lastSheetIndex < 0) {
			throw new IllegalArgumentException("Invalid sheetIndex2: " + lastSheetIndex + ".");
		}
		_bookEvaluator = bookEvaluator;
		_tracker = tracker;
		_sheetIndex = sheetIndex;
		_lastSheetIndex = lastSheetIndex;
	}

	public String getSheetName() {
		return _bookEvaluator.getSheetName(_sheetIndex);
	}
	
	public String getLastSheetName() {
		return _bookEvaluator.getSheetName(_lastSheetIndex);
	}

	public ValueEval getEvalForCell(int rowIndex, int columnIndex) {
		return _bookEvaluator.evaluateReference(getSheetName(), getLastSheetName(), rowIndex, columnIndex, _tracker);
	}
	
	public String getBookName() {
		final CollaboratingWorkbooksEnvironment env = _bookEvaluator.getEnvironment();
		return env.getBookName(_bookEvaluator);
	}
}
