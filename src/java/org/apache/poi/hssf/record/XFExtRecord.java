/* FullColorExt.java

	Purpose:
		
	Description:
		
	History:
		Jan 17, 2011 10:09:12 AM     2011, Created by henrichen

Copyright (C) 2011 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.poi.hssf.record;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.poi.hssf.record.common.FtrHeader;
import org.zkoss.poi.util.HexDump;
import org.zkoss.poi.util.LittleEndianOutput;

/**
 * 2.4.355 XFEXT [MS-XLS].pdf, page 609.
 */
public final class XFExtRecord extends StandardRecord  {
	public final static short sid = 0x087D;
	
	private FtrHeader futureHeader;
	
	private short reserved1; // Should always be zero
	private short ixfe;
	private short reserved2; // Should always be zero
	private int cexts;
	private Map<Short, ExtProp> rgexts;
	
	public XFExtRecord() {
		futureHeader = new FtrHeader();
		futureHeader.setRecordType(sid);
		cexts = 0;
		reserved1 = reserved2 = 0;
		rgexts = new HashMap<Short, ExtProp>();
	}

	public short getSid() {
		return sid;
	}

	public XFExtRecord(RecordInputStream in) {
		futureHeader = new FtrHeader(in);
		reserved1 = in.readShort();
		ixfe = (short) in.readUShort();
		reserved2 = in.readShort();
		cexts = in.readUShort();
		rgexts = new HashMap<Short, ExtProp>(cexts);
		for(int j= 0; j < cexts; ++j) {
			addExtProp(in);
		}
	}

	private void addExtProp(RecordInputStream in) {
		final ExtProp extProp = new ExtProp(in);
		rgexts.put(Short.valueOf(extProp.getExtType()), extProp);
	}
	
	public String toString() {
		StringBuffer sb = new StringBuffer();
		sb.append("[XFEXT]\n");
		sb.append("  .ixfe =").append(ixfe).append("\n");
		sb.append("  .cexts=").append(cexts).append("\n");
		for(ExtProp extProp : rgexts.values()) {
			sb.append("  .ExtProps\n");
			extProp.appendString(sb, "    ");
		}
		sb.append("[/XFEXT]\n");
		return sb.toString();
	}

	public void serialize(LittleEndianOutput out) {
		futureHeader.serialize(out);
		
		out.writeShort(reserved1);
		out.writeShort(ixfe);
		out.writeShort(reserved2);
		out.writeShort((short)cexts);
		for(ExtProp extProp : rgexts.values()) {
			extProp.serialize(out);
		}
	}

	public void cloneXFExtFrom(XFExtRecord source) {
		cexts = source.cexts;
		rgexts.putAll(source.rgexts);
	}
	
	protected int getDataSize() {
		int result = 12 + 2+2+2+2;
		for(ExtProp extProp : rgexts.values()) {
			result += extProp.getDataSize();
		}
		return result;
	}

	public int getIxfe() {
		return ixfe;
	}
	public void setIxfe(int ixfe) {
		this.ixfe = (short) ixfe;
	}
	public int getCexts() {
		return cexts;
	}
	public void setCexts(int cexts) {
		this.cexts = cexts;
	}
    
    //HACK: do a "cheat" clone, see Record.java for more information
    public Object clone() {
        return cloneViaReserialise();
    }
    
    public FullColorExt getFillForegroundColor() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x0004));
    	return getFullColorExt(extProp != null ? (FullColorExt) extProp.getExtPropData() : null);
    }
    
    public void setFillForegroundColor(FullColorExt color) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x0004))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x0004, color));
    }
    private void addExtProp(ExtProp extProp) {
    	rgexts.put(extProp.getExtType(), extProp);
    }
    public FullColorExt getFillBackgroundColor() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x0005));
    	return getFullColorExt(extProp != null ? (FullColorExt) extProp.getExtPropData() : null);
    }
    public void setFillBackgroundColor(FullColorExt color) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x0005))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x0005, color));
    }
    public XFExtGradient getGradientFill() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x0006));
    	return extProp != null ? (XFExtGradient) extProp.getExtPropData() : null;
    }
    public void setFillGradientFill(XFExtGradient color) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x0006))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x0006, color));
    }
    public FullColorExt getTopBorderColor() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x0007));
    	return getFullColorExt(extProp != null ? (FullColorExt) extProp.getExtPropData() : null);
    }
    public void setTopBorderColor(FullColorExt color) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x0007))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x0007, color));
    }
    public FullColorExt getBottomBorderColor() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x0008));
    	return getFullColorExt(extProp != null ? (FullColorExt) extProp.getExtPropData() : null);
    }
    public void setBottomBorderColor(FullColorExt color) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x0008))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x0008, color));
    }
    public FullColorExt getLeftBorderColor() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x0009));
    	return getFullColorExt(extProp != null ? (FullColorExt) extProp.getExtPropData() : null);
    }
    public void setLeftBorderColor(FullColorExt color) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x0009))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x0009, color));
    }
    public FullColorExt getRightBorderColor() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x000A));
    	return getFullColorExt(extProp != null ? (FullColorExt) extProp.getExtPropData() : null);
    }
    public void setRightBorderColor(FullColorExt color) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x000A))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x000A, color));
    }
    public FullColorExt getDiagonalBorderColor() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x000B));
    	return getFullColorExt(extProp != null ? (FullColorExt) extProp.getExtPropData() : null);
    }
    public void setDiagonalBorderColor(FullColorExt color) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x000B))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x000B, color));
    }
    public FullColorExt getTextColor() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x000D));
    	return getFullColorExt(extProp != null ? (FullColorExt) extProp.getExtPropData() : null);
    }
    public void setTextColor(FullColorExt color) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x000D))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x000D, color));
    }
    private FullColorExt getFullColorExt(FullColorExt ext) {
    	if (ext == null) {
    		return null;
    	}
    	if (ext.getXclrType() == 0 || ext.getXclrType() == 4) {
    		return null;
    	}
    	return ext;
    }
    public Byte getFontScheme() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x000E));
    	return extProp != null ? (Byte)extProp.getExtPropData() : null;
    }
    public void setFontScheme(Byte scheme) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x000E))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x000E, scheme));
    }
    public Byte getIndentLevel() {
    	final ExtProp extProp = rgexts.get(Short.valueOf((short)0x000F));
    	return extProp != null ? (Byte)extProp.getExtPropData() : null;
    }
    public void setIndentLevel(Byte indent) {
    	if (!rgexts.containsKey(Short.valueOf((short)0x000F))) {
        	++cexts;
    	}
    	addExtProp(new ExtProp((short)0x000F, indent));
    }
    private static class ExtProp {
    	private short extType;
    	private int cb;
    	private Object extPropData;
    	
    	private ExtProp(short extType, Object extPropData) {
    		this.extType = extType;
    		this.extPropData = extPropData;
    		cb = 4;
    		if (extPropData instanceof FullColorExt) {
    			cb += ((FullColorExt)extPropData).getDataSize();
    		} else if (extPropData instanceof XFExtGradient) {
    			cb += ((XFExtGradient)extPropData).getDataSize();
    		} else if (extPropData instanceof Byte) {
    			cb += 1;
    		} else {
    			throw new RuntimeException("Unknown extPropData. extType:"+extType+", extPropData:"+extPropData);
    		}
    	}
    	public void appendString(StringBuffer sb, String prefix) {
    		sb.append(prefix).append(".extType=").append(HexDump.shortToHex(extType)).append("\n");
    		sb.append(prefix).append(".cb     =").append(cb).append("\n");
    		sb.append(prefix).append(".extPropData\n");
    		if (extPropData instanceof FullColorExt) {
    			((FullColorExt)extPropData).appendString(sb, prefix+"  ");
    		} else if (extPropData instanceof XFExtGradient) {
    			((XFExtGradient)extPropData).appendString(sb, prefix+"  ");
    		} else if (extPropData instanceof Byte) {
    			sb.append(prefix+"  ").append(HexDump.byteToHex(((Byte)extPropData).byteValue())).append("\n");
    		} else {
    			sb.append(prefix+"  ").append("Unknown extPropData:"+extPropData).append("\n");
    		}
    	}	
    	
    	public ExtProp(RecordInputStream in) {
    		extType = in.readShort();
    		cb = in.readUShort();
    		switch(extType) {
    		case 0x0004:
    		case 0x0005:
    		case 0x0007:
    		case 0x0008:
    		case 0x0009:
    		case 0x000A:
    		case 0x000B:
    		case 0x000D:
    			extPropData = new FullColorExt(in);
    			break;
    		case 0x0006:
    			extPropData = new XFExtGradient(in);
    			break;
    		case 0x000E:
    		case 0x000F:
    			extPropData = Byte.valueOf(in.readByte()); //20110118, henrichen@zkoss.org: [MS-XLS].pdf say 2 bytes, but really only one byte
    			break;
    		default:
    			throw new RuntimeException("Unknown extType:"+extType);
    		}
    	}
    	
    	private int getDataSize() {
    		return cb;
    	}
    	
    	private short getExtType() {
    		return extType;
    	}
    	
    	private Object getExtPropData() {
    		return extPropData;
    	}
    	
    	private void serialize(LittleEndianOutput out) {
    		out.writeShort(extType);
    		out.writeShort(cb);
    		if (extPropData instanceof FullColorExt) {
    			((FullColorExt)extPropData).serialize(out);
    		} else if (extPropData instanceof XFExtGradient) {
    			((XFExtGradient)extPropData).serialize(out);
    		} else if (extPropData instanceof Byte) {
    			out.writeByte(((Byte)extPropData).byteValue());
    		} else {
    			throw new RuntimeException("Unknown extPropData:"+extPropData);
    		}
    	}
    }
    
    //TODO not complete yet
    private static class XFExtGradient {
    	private XFPropGradient gradient;
    	private int cGradStops; //0~256
    	private GradStop[] rgGradStops;
    	
    	public XFExtGradient(RecordInputStream in) {
    		gradient = new XFPropGradient(in);
    		cGradStops = in.readInt();
    		rgGradStops = new GradStop[cGradStops];
    		for(int j = 0; j < cGradStops; ++j) {
    			rgGradStops[j] = new GradStop(in);
    		}
    	}
    	public int getDataSize() {
    		return gradient.getDataSize() + 4 + cGradStops * GradStop.DATA_SIZE;
    	}
    	public void appendString(StringBuffer sb, String prefix) {
    		sb.append(prefix).append(".gradient\n");
    		gradient.toString(sb, prefix+"  ");
    		sb.append(prefix).append(".cGradStops=").append(cGradStops).append("\n");
    		sb.append(prefix).append(".rgGradStops\n");
    		for(int j = 0; j < cGradStops; ++j) {
    			rgGradStops[j].toString(sb, prefix+"  ");
    		}
    	}
    	public void serialize(LittleEndianOutput out) {
    		gradient.serialize(out);
    		out.writeInt(cGradStops);
    		for(int j = 0; j < cGradStops; ++j) {
    			rgGradStops[j].serialize(out);
    		}
    	}
    	
    }
    
    //TODO not complete yet
    private static class XFPropGradient {
    	private static final int DATA_SIZE = 4 + 8 + 8 + 8 + 8 + 8; 
    	private int type;
    	private double numDegree;
    	private double numFillToLeft;
    	private double numFillToRight;
    	private double numFillToTop;
    	private double numFillToBottom;
    	
    	public XFPropGradient(RecordInputStream in) {
    		type = in.readInt();
    		numDegree = in.readDouble();
    		numFillToLeft = in.readDouble();
    		numFillToRight = in.readDouble();
    		numFillToTop = in.readDouble();
    		numFillToBottom = in.readDouble();
    	}
    	public int getDataSize() {
    		return DATA_SIZE;
    	}
    	public void toString(StringBuffer sb, String prefix) {
    		sb.append(prefix).append(".type          =").append(type).append("\n");
    		sb.append(prefix).append(".numDegree     =").append(numDegree).append("\n");
    		sb.append(prefix).append(".numFillToLeft =").append(numFillToLeft).append("\n");
    		sb.append(prefix).append(".numFillToRight=").append(numFillToRight).append("\n");
    		sb.append(prefix).append(".numFillToTop  =").append(numFillToTop).append("\n");
    		sb.append(prefix).append(".numFillToBottom=").append(numFillToBottom).append("\n");
    	}
    	public void serialize(LittleEndianOutput out) {
    		out.writeInt(type);
    		out.writeDouble(numDegree);
    		out.writeDouble(numFillToLeft);
    		out.writeDouble(numFillToRight);
    		out.writeDouble(numFillToTop);
    		out.writeDouble(numFillToBottom);
    	}
    }
    
    //2.5.156 GradStop [MS-XLS].pdf, page 731
    //TODO not complete yet
    private static class GradStop {
    	private static final int DATA_SIZE = 2 + 4 + 8 + 8; 
    	private short xclrType;
    	private int xclrValue;
    	private double numPosition;
    	private double numTint;
    	
    	public GradStop(RecordInputStream in) {
    		xclrType = in.readShort();
    		xclrValue = in.readInt();
    		numPosition = in.readDouble();
    		numTint = in.readDouble();
    	}
    	public void toString(StringBuffer sb, String prefix) {
    		sb.append(prefix).append(".xclrType   =").append(xclrType).append("\n");
    		sb.append(prefix).append(".xclrValue  =").append(xclrValue).append("\n");
    		sb.append(prefix).append(".numPosition=").append(numPosition).append("\n");
    		sb.append(prefix).append(".numTint    =").append(numTint).append("\n");
    	}
    	public void serialize(LittleEndianOutput out) {
    		out.writeShort(xclrType);
    		out.writeInt(xclrValue);
    		out.writeDouble(numPosition);
    		out.writeDouble(numTint);
    	}
    	public int getDataSize() {
    		return DATA_SIZE;
    	}
    }
}
