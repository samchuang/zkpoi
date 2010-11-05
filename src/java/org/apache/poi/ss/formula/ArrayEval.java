/* ArrayEval.java

	Purpose:
		
	Description:
		
	History:
		Oct 21, 2010 2:19:12 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.poi.ss.formula;

import org.zkoss.poi.hssf.record.formula.AreaI;
import org.zkoss.poi.hssf.record.formula.ArrayPtg;
import org.zkoss.poi.hssf.record.formula.eval.AreaEval;
import org.zkoss.poi.hssf.record.formula.eval.NumberEval;
import org.zkoss.poi.hssf.record.formula.eval.StringEval;
import org.zkoss.poi.hssf.record.formula.eval.ValueEval;

import com.sun.org.apache.xml.internal.utils.UnImplNode;

/**
 * Constant value array eval.
 * @author henrichen
 *
 */
public class ArrayEval implements AreaEval {
	private final int _nColumns;
	private final int _nRows;
	private final Object[][] _values;

	protected ArrayEval(Object[][] srcvalues, int firstRow, int firstColumn, int lastRow, int lastColumn) {
		_nColumns = lastColumn - firstColumn + 1;
		_nRows = lastRow - firstRow + 1;
		_values = new Object[_nRows][];
		for(int r = firstRow, j = 0; r <= lastRow; ++r) {
			final Object[] dst = new Object[_nColumns]; 
			_values[j++] = dst;
			System.arraycopy(srcvalues[r], firstColumn, dst, 0, _nColumns);
		}
	}
	
	protected ArrayEval(ArrayPtg ptg) {
		_nColumns = ptg.getColumnCount();
		_nRows = ptg.getRowCount();
		_values = ptg.getTokenArrayValues();
	}

	@Override
	public boolean contains(int row, int col) {
		return 0 <= row && row < _nRows && 0 <= col && col < _nColumns;
	}

	@Override
	public boolean containsColumn(int col) {
		return 0 <= col && col < _nColumns;
	}

	@Override
	public boolean containsRow(int row) {
		return 0 <= row && row < _nRows;
	}

	@Override
	public ValueEval getAbsoluteValue(int row, int col) {
		return getRelativeValue(row, col);
	}

	@Override
	public int getFirstColumn() {
		return 0;
	}

	@Override
	public int getFirstRow() {
		return 0;
	}

	@Override
	public int getHeight() {
		return _nRows;
	}

	@Override
	public int getLastColumn() {
		return _nColumns - 1;
	}

	@Override
	public int getLastRow() {
		return _nRows - 1;
	}

	@Override
	public ValueEval getRelativeValue(int row, int col) {
		if(!containsRow(row)) {
			throw new IllegalArgumentException("Specified row index (" + row
					+ ") is outside the allowed range (" + getFirstRow() + ".." + getLastRow() + ")");
		}
		if(!containsColumn(col)) {
			throw new IllegalArgumentException("Specified column index (" + col
					+ ") is outside the allowed range (" + getFirstColumn() + ".." + getLastColumn() + ")");
		}
		final Object val = _values[row][col];
		if (val instanceof Number) {
			return new NumberEval(((Number)val).doubleValue());
		} else if (val instanceof String) {
			return new StringEval((String) val);
		} else {
			throw new RuntimeException("Unknown Object type in ArrayPtg:" + val);
		}
	}

	@Override
	public int getWidth() {
		return _nColumns;
	}

	@Override
	public AreaEval offset(int relFirstRowIx, int relLastRowIx, int relFirstColIx, int relLastColIx) {
		return this;
	}

	@Override
	public TwoDEval getColumn(int columnIndex) {
		return new ArrayEval(_values, getFirstRow(), columnIndex, getLastRow(), columnIndex);
	}

	@Override
	public TwoDEval getRow(int rowIndex) {
		return new ArrayEval(_values, rowIndex, getFirstColumn(), rowIndex, getLastColumn());
	}

	@Override
	public ValueEval getValue(int rowIndex, int columnIndex) {
		return getRelativeValue(rowIndex, columnIndex);
	}

	@Override
	public boolean isColumn() {
		return false;
	}

	@Override
	public boolean isRow() {
		return false;
	}
}
