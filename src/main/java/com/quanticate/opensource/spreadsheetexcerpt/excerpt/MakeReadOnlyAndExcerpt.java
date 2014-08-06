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
package com.quanticate.opensource.spreadsheetexcerpt.excerpt;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import org.alfresco.service.cmr.repository.ContentReader;
import org.alfresco.service.cmr.repository.ContentWriter;

import com.quanticate.opensource.spreadsheetexcerpt.SpreadsheetHandler;

public interface MakeReadOnlyAndExcerpt extends SpreadsheetHandler
{
   /**
    * Marks the sheets to keep as read only (removing formulas), and
    *  removes all other sheets
    */
   public void excerpt(String[] sheetsToKeep, File input, OutputStream output) throws IOException;

   /**
    * Marks the sheets to keep as read only (removing formulas), and
    *  removes all other sheets.
    */
   public void excerpt(String[] sheetsToKeep, ContentReader input, ContentWriter output) throws IOException;

   /**
    * Marks the sheets to keep as read only (removing formulas), and
    *  removes all other sheets
    */
   public void excerpt(int[] sheetsToKeep, File input, OutputStream output) throws IOException;

   /**
    * Marks the sheets to keep as read only (removing formulas), and
    *  removes all other sheets
    */
   public void excerpt(int[] sheetsToKeep, ContentReader input, ContentWriter output) throws IOException;
}
