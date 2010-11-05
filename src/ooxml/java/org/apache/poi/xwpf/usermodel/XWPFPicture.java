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
package org.zkoss.poi.xwpf.usermodel;

import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.zkoss.poi.POIXMLDocumentPart;
import org.zkoss.poi.openxml4j.opc.PackageRelationship;
import org.zkoss.poi.util.POILogFactory;
import org.zkoss.poi.util.POILogger;


/**
 * @author Philipp Epp
 */
public class XWPFPicture {
	private static final POILogger logger = POILogFactory.getLogger(XWPFPicture.class);
	
	protected XWPFParagraph paragraph;
	private CTPicture ctPic;
	 
	 public XWPFParagraph getParagraph(){
		 return paragraph;
	 }
	 
	 public XWPFPicture(CTPicture ctPic, XWPFParagraph paragraph){
		 this.paragraph = paragraph;
		 this.ctPic = ctPic;
	 }
	 
	 /**
	  * Link Picture with PictureData
	  * @param rel
	  */
	 public void setPictureReference(PackageRelationship rel){
		 ctPic.getBlipFill().getBlip().setEmbed(rel.getId());
	 }
	 
    /**
     * Return the underlying CTPicture bean that holds all properties for this picture
     *
     * @return the underlying CTPicture bean
     */
    public CTPicture getCTPicture(){
        return ctPic;
    }
    
    /**
     * Get the PictureData of the Picture, if present.
     * Note - not all kinds of picture have data
     */
    public XWPFPictureData getPictureData(){
    	String blipId = ctPic.getBlipFill().getBlip().getEmbed();
    	for(POIXMLDocumentPart part: paragraph.getDocument().getRelations()){
    		  if(part.getPackageRelationship().getId().equals(blipId)){
    			  return (XWPFPictureData)part;
    		  }
    	}
    	return null;
    }
    
}
