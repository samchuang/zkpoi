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
 * Interface for constructing the formula dependency.
 *
 * @author Henri Chen (henrichen at zkoss dot org) - dependency tracking
 */
public interface DependencyTracker {
	/**
	 * Callback when evaluating a formula cell.
	 * 
	 * @param ec the evaluation context of the evaluated formula cell.
	 * @param opResult the precedent that might change the formula cell.
	 * @param eval whether this reference is an evaluated result(e.g. from INDIRECT() function(true), or directly specified in formula(false))
	 * @return the ValueEval after the dependency checking
	 */
	public ValueEval addDependency(OperationEvaluationContext ec, ValueEval opResult, boolean eval);
}
