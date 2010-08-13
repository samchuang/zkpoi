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

package org.apache.poi.ss.util;


/**
 * Helper methods for when working with Usermodel Workbooks
 */
public class WorkbookUtil {
	
	/**
	 * Creates a valid sheet name, which is conform to the rules.
	 * In any case, the result safely can be used for 
	 * {@link org.apache.poi.hssf.usermodel.HSSFWorkbook#setSheetName(int, String)}.
	 * <br>
	 * Rules:
	 * <ul>
	 * <li>never null</li>
	 * <li>minimum length is 1</li>
	 * <li>maximum length is 31</li>
	 * <li>doesn't contain special chars: / \ ? * ] [ </li>
	 * <li>Sheet names must not begin or end with ' (apostrophe)</li>
	 * </ul>
	 * Invalid characters are replaced by one space character ' '.
	 * 
	 * @param nameProposal can be any string, will be truncated if necessary,
	 *        allowed to be null
	 * @return a valid string, "empty" if to short, "null" if null         
	 */
	public final static String createSafeSheetName(final String nameProposal) {
		if (nameProposal == null) {
			return "null";
		}
		if (nameProposal.length() < 1) {
			return "empty";
		}
		final int length = Math.min(31, nameProposal.length());
		final String shortenname = nameProposal.substring(0, length);
		final StringBuilder result = new StringBuilder(shortenname);
		for (int i=0; i<length; i++) {
			char ch = result.charAt(i);
			switch (ch) {
				case '/':
				case '\\':
				case '?':
				case '*':
				case ']':
				case '[':
					result.setCharAt(i, ' ');
					break;
				case '\'':
					if (i==0 || i==length-1) {
						result.setCharAt(i, ' ');
					}
					break;
				default:
					// all other chars OK
			}
		}
		return result.toString();
	}
}
