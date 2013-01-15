package com.quanticate.opensource.spreadsheetexcerpt.webscript;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.springframework.extensions.webscripts.Cache;
import org.springframework.extensions.webscripts.DeclarativeWebScript;
import org.springframework.extensions.webscripts.Status;
import org.springframework.extensions.webscripts.WebScriptRequest;

import com.quanticate.opensource.spreadsheetexcerpt.excerpt.MakeReadOnlyAndExcerpt;

// TODO Should this be AbstractWebScript instead?
public class SpreadsheetExcerptPost extends DeclarativeWebScript
{
   private MakeReadOnlyAndExcerpt excerpter;

   @Override
   protected Map<String, Object> executeImpl(WebScriptRequest req,
         Status status, Cache cache)
   {
      // TODO
      return new HashMap<String, Object>();
   }

   public void setExcerpter(MakeReadOnlyAndExcerpt excerpter)
   {
      this.excerpter = excerpter;
   }
}
