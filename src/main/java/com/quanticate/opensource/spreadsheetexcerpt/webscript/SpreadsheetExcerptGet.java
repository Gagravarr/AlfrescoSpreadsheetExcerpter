package com.quanticate.opensource.spreadsheetexcerpt.webscript;

import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.extensions.webscripts.Cache;
import org.springframework.extensions.webscripts.DeclarativeWebScript;
import org.springframework.extensions.webscripts.Status;
import org.springframework.extensions.webscripts.WebScriptRequest;

import com.quanticate.opensource.spreadsheetexcerpt.excerpt.MakeReadOnlyAndExcerpt;

public class SpreadsheetExcerptGet extends DeclarativeWebScript
{
   private MakeReadOnlyAndExcerpt excerpter;

   @Override
   protected Map<String, Object> executeImpl(WebScriptRequest req,
         Status status, Cache cache)
   {
      Map<String,Object> model = new HashMap<String, Object>();
      model.put("file", "TODO");
      model.put("sheets", Collections.EMPTY_LIST);
      return model;
   }

   public void setExcerpter(MakeReadOnlyAndExcerpt excerpter)
   {
      this.excerpter = excerpter;
   }
}
