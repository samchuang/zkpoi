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
package org.zkoss.poi.hssf.util;

import org.zkoss.poi.hssf.record.FullColorExt;
import org.zkoss.poi.util.HexDump;
/**
 * HSSFColor that wrap FullColorExt.
 * @author henrichen
 */
public final class HSSFColorExt extends HSSFColor
{
	private final FullColorExt _colorExt;

	public HSSFColorExt(FullColorExt colorExt) {
		_colorExt = colorExt;
	}
	
    public short getIndex()
    {
        return _colorExt.isIndex() ? (short)_colorExt.getXclrValue() : AUTOMATIC.getInstance().getIndex();
    }

    public boolean isIndex() {
    	return _colorExt.isIndex();
    }
    
    public short [] getTriplet()
    {
    	if (_colorExt.isRGB()) {
    		final int color = _colorExt.getXclrValue();
    		final short[] rgb = new short[3];
    		rgb[0] = (short) ((color >> 16) & 0xff); //r
    		rgb[1] = (short) ((color >> 8) & 0xff); //g
    		rgb[2] = (short) (color & 0xff); //b
            return rgb;
    	}
    	return AUTOMATIC.getInstance().getTriplet();
    }

    public String getHexString()
    {
    	final short[] rgb = getTriplet();
    	if (rgb != null) {
        	String r = new String(HexDump.byteToHex(rgb[0]));
        	String g = new String(HexDump.byteToHex(rgb[1]));
        	String b = new String(HexDump.byteToHex(rgb[2]));
            return (rgb[0] == 0 ? "0" : (r+r)) + ":"
            		+ (rgb[1] == 0 ? "0" : (g+g)) + ":"
            		+ (rgb[2] == 0 ? "0" : (b+b));
    	}
    	return AUTOMATIC.getInstance().getHexString();
    }
}
