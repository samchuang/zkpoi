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

package org.apache.poi.hsmf;

import java.io.IOException;

import junit.framework.TestCase;

import org.apache.poi.POIDataSamples;
import org.apache.poi.hsmf.exceptions.ChunkNotFoundException;

/**
 * Tests to verify that we can perform basic opperations on 
 *  a range of files
 */
public final class TestBasics extends TestCase {
   private MAPIMessage simple;
   private MAPIMessage quick;
   private MAPIMessage outlook30;
   private MAPIMessage attachments;

	/**
	 * Initialize this test, load up the blank.msg mapi message.
	 * @throws Exception
	 */
	public TestBasics() throws IOException {
        POIDataSamples samples = POIDataSamples.getHSMFInstance();
		simple = new MAPIMessage(samples.openResourceAsStream("simple_test_msg.msg"));
      quick  = new MAPIMessage(samples.openResourceAsStream("quick.msg"));
      outlook30  = new MAPIMessage(samples.openResourceAsStream("outlook_30_msg.msg"));
      attachments = new MAPIMessage(samples.openResourceAsStream("attachment_test_msg.msg"));
	}
	
	/**
	 * Can we always get the recipient's email?
	 */
	public void testRecipientEmail() throws Exception {
      assertEquals("travis@overwrittenstack.com", simple.getRecipientEmailAddress());
      assertEquals("kevin.roast@alfresco.org", quick.getRecipientEmailAddress());
      assertEquals("nicolas1.23456@free.fr", attachments.getRecipientEmailAddress());
      
      // This one has lots...
      assertEquals(18, outlook30.getRecipientEmailAddressList().length);
      assertEquals("shawn.bohn@pnl.gov; gus.calapristi@pnl.gov; Richard.Carter@pnl.gov; " +
      		"barb.cheney@pnl.gov; nick.cramer@pnl.gov; vern.crow@pnl.gov; Laura.Curtis@pnl.gov; " +
      		"julie.dunkle@pnl.gov; david.gillen@pnl.gov; michelle@pnl.gov; Jereme.Haack@pnl.gov; " +
      		"Michelle.Hart@pnl.gov; ranata.johnson@pnl.gov; grant.nakamura@pnl.gov; " +
      		"debbie.payne@pnl.gov; stuart.rose@pnl.gov; randall.scarberry@pnl.gov; Leigh.Williams@pnl.gov", 
            outlook30.getRecipientEmailAddress()
      );
	}
	
	/**
	 * Test subject
	 */
	public void testSubject() throws Exception {
      assertEquals("test message", simple.getSubject());
      assertEquals("Test the content transformer", quick.getSubject());
      assertEquals("IN-SPIRE servers going down for a bit, back up around 8am", outlook30.getSubject());
      assertEquals("test pi\u00e8ce jointe 1", attachments.getSubject());
	}
	
	/**
	 * Test attachments
	 */
	public void testAttachments() throws Exception {
      assertEquals(0, simple.getAttachmentFiles().length);
      assertEquals(0, quick.getAttachmentFiles().length);
      assertEquals(0, outlook30.getAttachmentFiles().length);
      assertEquals(2, attachments.getAttachmentFiles().length);
	}
	
	/**
	 * Test missing chunks
	 */
	public void testMissingChunks() throws Exception {
	   assertEquals(false, attachments.isReturnNullOnMissingChunk());

	   try {
	      attachments.getMessageDate();
	      fail();
	   } catch(ChunkNotFoundException e) {
	      // Good
	   }
	   
	   attachments.setReturnNullOnMissingChunk(true);
	   
	   assertEquals(null, attachments.getMessageDate());
	   
      attachments.setReturnNullOnMissingChunk(false);
      
      try {
         attachments.getMessageDate();
         fail();
      } catch(ChunkNotFoundException e) {
         // Good
      }
	}
}
