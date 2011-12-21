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
package org.apache.poi.xssf.model;

import java.io.IOException;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTColorScheme;
import org.openxmlformats.schemas.drawingml.x2006.main.ThemeDocument;
import org.openxmlformats.schemas.drawingml.x2006.main.CTColor;

/**
 * Class that represents theme of XLSX document. The theme includes specific
 * colors and fonts.
 * 
 * @author Petr Udalau(Petr.Udalau at exigenservices.com) - theme colors
 */
public class ThemesTable extends POIXMLDocumentPart {
    private ThemeDocument theme;

    public ThemesTable(PackagePart part, PackageRelationship rel) throws IOException {
        super(part, rel);
        
        try {
           theme = ThemeDocument.Factory.parse(part.getInputStream());
        } catch(XmlException e) {
           throw new IOException(e.getLocalizedMessage());
        }
    }

    public ThemesTable(ThemeDocument theme) {
        this.theme = theme;
    }

    public XSSFColor getThemeColor(int idx) {
        CTColorScheme colorScheme = theme.getTheme().getThemeElements().getClrScheme();
        CTColor ctColor = null;
        int cnt = 0;
        for (XmlObject obj : colorScheme.selectPath("./*")) {
            if (obj instanceof org.openxmlformats.schemas.drawingml.x2006.main.CTColor) {
                if (cnt == idx) {
                    ctColor = (org.openxmlformats.schemas.drawingml.x2006.main.CTColor) obj;
                    
                    byte[] rgb = null;
                    if (ctColor.getSrgbClr() != null) {
                       // Colour is a regular one 
                       rgb = ctColor.getSrgbClr().getVal();
                    } else if (ctColor.getSysClr() != null) {
                       // Colour is a tint of white or black
                       rgb = ctColor.getSysClr().getLastClr();
                    }

                    return new XSSFColor(rgb);
                }
                cnt++;
            }
        }
        return null;
    }
    
    /**
     * If the colour is based on a theme, then inherit 
     *  information (currently just colours) from it as
     *  required.
     */
    public void inheritFromThemeAsRequired(XSSFColor color) {
       if(color == null) {
          // Nothing for us to do
          return;
       }
       if(! color.getCTColor().isSetTheme()) {
          // No theme set, nothing to do
          return;
       }
       
       // Get the theme colour
       XSSFColor themeColor = getThemeColor(color.getTheme());
       // Set the raw colour, not the adjusted one
       // Do a raw set, no adjusting at the XSSFColor layer either
       color.getCTColor().setRgb(themeColor.getCTColor().getRgb());
       
       // All done
    }
}
