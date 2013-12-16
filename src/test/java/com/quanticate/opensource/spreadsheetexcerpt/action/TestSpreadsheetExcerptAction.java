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
