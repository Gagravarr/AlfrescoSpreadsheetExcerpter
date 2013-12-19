/* ====================================================================
  Copyright 2013 Quanticate Ltd

  Licensed under the Apache License, Version 2.0 (the "License");
  you may not use this file except in compliance with the License.
  You may obtain a copy of the License at

      http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
==================================================================== */
package com.quanticate.opensource.spreadsheetexcerpt.action;

import org.alfresco.util.test.junitrules.ApplicationContextInit;
import org.junit.BeforeClass;
import org.junit.ClassRule;
import org.junit.Test;

public class TestSpreadsheetExcerptAction
{
   private static SpreadsheetExcerptActionExecutor action;

   @ClassRule public static ApplicationContextInit APP_CONTEXT_INIT = 
         new ApplicationContextInit();
   
   @BeforeClass
   public static void setup()
   {
      // Get the action executor
      action = APP_CONTEXT_INIT.getApplicationContext().getBean(
            "action.spreadsheetexcerpt", SpreadsheetExcerptActionExecutor.class);
      
      // Create some sample files
      // TODO
   }
   
   @Test
   public void invalidParameters() throws Exception
   {
      // TODO
   }
   
   @Test
   public void basicExcerptXLS() throws Exception
   {
      // TODO
   }
   
   @Test
   public void basicExcerptXLSX() throws Exception
   {
      // TODO
   }
   
   @Test
   public void allParameters() throws Exception
   {
      // TODO
   }
}
