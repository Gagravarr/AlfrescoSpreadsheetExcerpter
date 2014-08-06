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
package com.quanticate.opensource.spreadsheetexcerpt.merge;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

import org.alfresco.service.cmr.repository.ContentReader;
import org.alfresco.service.cmr.repository.ContentWriter;

import com.quanticate.opensource.spreadsheetexcerpt.SpreadsheetHandler;
import com.quanticate.opensource.spreadsheetexcerpt.excerpt.MakeReadOnlyAndExcerpt;

/**
 * Opposite of {@link MakeReadOnlyAndExcerpt} - takes values from
 *  an Excerpted sheet and merges them back into a master one
 */
public interface MergeChangesFromExcerpt extends SpreadsheetHandler
{
   /**
    * For matching Sheets in the two files, for non-Formula cells in the
    *  Full version, copy over the (possibly) updated values from the
    *  Excerpt version.
    *
    * @param sheetsToMerge The names of the sheets to merge from the excerpt file
    * @param excerptInput The file to read the cells to merge from
    * @param fullInput The full file, which will have data in suitable cells in the given sheets replaced
    * @param output Where to write the merged file to
    */
   public void merge(String[] sheetsToMerge, File excerptInput, File fullInput, OutputStream output) throws IOException;

   /**
    * For matching Sheets in the two files, for non-Formula cells in the
    *  Full version, copy over the (possibly) updated values from the
    *  Excerpt version.
    *
    * @param sheetsToMerge The names of the sheets to merge from the excerpt file
    * @param excerptInput The source file to read the cells to merge from
    * @param fullInput The full source file, which will have data in suitable cells in the given sheets replaced
    * @param output Where to write the merged file to
    */
   public void merge(String[] sheetsToMerge, ContentReader excerptInput, ContentReader fullInput, ContentWriter output) throws IOException;
}
