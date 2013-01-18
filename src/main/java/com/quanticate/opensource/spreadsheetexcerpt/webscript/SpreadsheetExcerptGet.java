package com.quanticate.opensource.spreadsheetexcerpt.webscript;

import java.util.HashMap;
import java.util.Map;

import org.springframework.extensions.webscripts.Cache;
import org.springframework.extensions.webscripts.DeclarativeWebScript;
import org.springframework.extensions.webscripts.Status;
import org.springframework.extensions.webscripts.WebScriptRequest;

import com.quanticate.opensource.spreadsheetexcerpt.excerpt.MakeReadOnlyAndExcerpt;

/**
 * Currently does nothing, all the work is done in
 *  {@link SpreadsheetExcerptPickPost} and 
 *  {@link SpreadsheetExcerptPost}
 */
public class SpreadsheetExcerptGet extends DeclarativeWebScript
{
   @SuppressWarnings("unused")
   private MakeReadOnlyAndExcerpt excerpter;

   @Override
   protected Map<String, Object> executeImpl(WebScriptRequest req,
         Status status, Cache cache)
   {
      Map<String,Object> model = new HashMap<String, Object>(0);
      return model;
   }

   public void setExcerpter(MakeReadOnlyAndExcerpt excerpter)
   {
      this.excerpter = excerpter;
   }
}
