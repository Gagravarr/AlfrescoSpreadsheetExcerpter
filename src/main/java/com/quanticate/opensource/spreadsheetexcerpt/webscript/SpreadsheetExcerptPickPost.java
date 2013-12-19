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
package com.quanticate.opensource.spreadsheetexcerpt.webscript;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.alfresco.util.TempFileProvider;
import org.apache.poi.util.IOUtils;
import org.springframework.extensions.webscripts.Cache;
import org.springframework.extensions.webscripts.DeclarativeWebScript;
import org.springframework.extensions.webscripts.Status;
import org.springframework.extensions.webscripts.WebScriptException;
import org.springframework.extensions.webscripts.WebScriptRequest;
import org.springframework.extensions.webscripts.servlet.FormData;

import com.quanticate.opensource.spreadsheetexcerpt.excerpt.MakeReadOnlyAndExcerpt;

/**
 * Places the spreadsheet in a medium lived temporary file,
 *  and returns the list of Sheet Names in it
 */
public class SpreadsheetExcerptPickPost extends DeclarativeWebScript
{
   private MakeReadOnlyAndExcerpt excerpter;

   @Override
   protected Map<String, Object> executeImpl(WebScriptRequest req,
         Status status, Cache cache)
   {
      // Check we got a valid Multi-Part Form Upload
      FormData form = (FormData)req.parseContent();
      if (form == null || !form.getIsMultiPart())
      {
          throw new WebScriptException(Status.STATUS_BAD_REQUEST, "Bad Upload");
      }
      
      // Look for a file, we're quite flexible!
      File spreadsheet = null;
      for (FormData.FormField field : form.getFields())
      {
          if (field.getIsFile())
          {
             String ext = ".xls";
             if (field.getFilename() != null && field.getFilename().endsWith(".xlsx"))
             {
                ext = ".xlsx";
             }
             
             try
             {
                spreadsheet = TempFileProvider.createTempFile("spreadsheet", ext);
                FileOutputStream sout = new FileOutputStream(spreadsheet);
                IOUtils.copy(field.getInputStream(), sout);
                sout.close();
             }
             catch (IOException e)
             {
                throw new WebScriptException(Status.STATUS_BAD_REQUEST, "Upload Failed");
             }
                   
             break;
          }
      }
      
      // Process
      String[] sheetnames = null;
      if (spreadsheet != null)
      {
         try
         {
            sheetnames = excerpter.getSheetNames(spreadsheet);
         }
         catch (IOException e)
         {
            throw new WebScriptException(Status.STATUS_BAD_REQUEST, "Spreadsheet Corrupt", e);
         }
      }
      else
      {
         throw new WebScriptException(Status.STATUS_BAD_REQUEST, "Spreadsheet Missing");
      }
      
      // Report
      Map<String,Object> model = new HashMap<String, Object>();
      model.put("file", spreadsheet.getName());
      model.put("sheets", sheetnames);
      return model;
   }

   public void setExcerpter(MakeReadOnlyAndExcerpt excerpter)
   {
      this.excerpter = excerpter;
   }
}
