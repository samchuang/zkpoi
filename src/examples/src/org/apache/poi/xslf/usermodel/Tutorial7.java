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

package org.apache.poi.xslf.usermodel;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Bullets and numbering
 *
 * @author Yegor Kozlov
 */
public class Tutorial7 {

    public static void main(String[] args) throws IOException{
        XMLSlideShow ppt = new XMLSlideShow();

        XSLFSlide slide = ppt.createSlide();
        XSLFTextBox shape = slide.createTextBox();
        shape.setAnchor(new Rectangle(50, 50, 400, 200));

        XSLFTextParagraph p1 = shape.addNewTextParagraph();
        p1.setLevel(0);
        p1.setBullet(true);
        XSLFTextRun r1 = p1.addNewTextRun();
        r1.setText("Bullet1");

        XSLFTextParagraph p2 = shape.addNewTextParagraph();
        // indentation before text
        p2.setLeftMargin(60);
        // the bullet is set 40 pt before the text
        p2.setIndent(-40);
        p2.setBullet(true);
        // customize bullets
        p2.setBulletFontColor(Color.red);
        p2.setBulletFont("Wingdings");
        p2.setBulletCharacter("\u0075");
        p2.setLevel(1);
        XSLFTextRun r2 = p2.addNewTextRun();
        r2.setText("Bullet2");

        // the next three paragraphs form an auto-numbered list
        XSLFTextParagraph p3 = shape.addNewTextParagraph();
        p3.setBulletAutoNumber(ListAutoNumber.ALPHA_LC_PARENT_R, 1);
        p3.setLevel(2);
        XSLFTextRun r3 = p3.addNewTextRun();
        r3.setText("Numbered List Item - 1");

        XSLFTextParagraph p4 = shape.addNewTextParagraph();
        p4.setBulletAutoNumber(ListAutoNumber.ALPHA_LC_PARENT_R, 2);
        p4.setLevel(2);
        XSLFTextRun r4 = p4.addNewTextRun();
        r4.setText("Numbered List Item - 2");

        XSLFTextParagraph p5 = shape.addNewTextParagraph();
        p5.setBulletAutoNumber(ListAutoNumber.ALPHA_LC_PARENT_R, 3);
        p5.setLevel(2);
        XSLFTextRun r5 = p5.addNewTextRun();
        r5.setText("Numbered List Item - 3");

        shape.resizeToFitText();

        FileOutputStream out = new FileOutputStream("list.pptx");
        ppt.write(out);
        out.close();
    }
}
