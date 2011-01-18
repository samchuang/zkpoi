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
package org.zkoss.poi.hssf.record;

import org.zkoss.poi.util.HexDump;
import org.zkoss.poi.util.LittleEndianOutput;

/**
 * 2.5.155 FullColorExt. ([MS-XLS].pdf, page 730.
 * @author henrichen
 *
 */
public class FullColorExt {
	private short xclrType; //2.5.279 XColorType. [MS-XLS].pdf, page 938.
	private short nTintShade;
	private int xclrValue;
	private byte[] unused = new byte[8];
	
	public FullColorExt(RecordInputStream in) {
		xclrType = in.readShort();
		nTintShade = in.readShort();
		xclrValue = in.readInt();
		in.read(unused, 0, unused.length);
	}
	public boolean isTheme() {
		return xclrType == 3;
	}
	public boolean isRGB() {
		return xclrType == 2;
	}
	public boolean isIndex() {
		return xclrType == 1;
	}
	public void appendString(StringBuffer sb, String prefix) {
		sb.append(prefix).append(".xclrType  =").append(xclrType).append("\n");
		sb.append(prefix).append(".nTintShade=").append(nTintShade).append("\n");
		sb.append(prefix).append(".xclrValue =").append(HexDump.intToHex(xclrValue)).append("\n");
	}
	public void serialize(LittleEndianOutput out) {
		out.writeShort(xclrType);
		out.writeShort(nTintShade);
		out.writeInt(xclrValue);
		out.write(unused);
	}
	
	public short getXclrType() {
		return xclrType;
	}
	
	public void setXclrType(short xclrType) {
		this.xclrType = xclrType;
	}
	
	public short getTintShade() {
		return nTintShade;
	}
	
	public void setTintShade(short nTintShade) {
		this.nTintShade = nTintShade;
	}
	
	public int getXclrValue() {
		return xclrValue;
	}
	
	public void setXclrValue(int xclrValue) {
		this.xclrValue = xclrValue;
	}
}
