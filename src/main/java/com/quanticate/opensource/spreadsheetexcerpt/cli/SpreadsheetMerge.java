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
package com.quanticate.opensource.spreadsheetexcerpt.cli;

import java.io.File;
import java.io.FileOutputStream;

import org.alfresco.util.Pair;

import com.quanticate.opensource.spreadsheetexcerpt.excerpt.POIExcerpterAndMerger;
import com.quanticate.opensource.spreadsheetexcerpt.merge.MergeChangesFromExcerpt;

/**
 * CLI Tool for running a merge
 */
public class SpreadsheetMerge extends SpreadsheetCLI
{
   public static void main(String[] args) throws Exception
   {
      Pair<File[],int[]> opts = processArgs(args, 2, "SpreadsheetMerge");
      
      MergeChangesFromExcerpt merger = new POIExcerpterAndMerger();
      
      File excerpt = opts.getFirst()[0];
      File full = opts.getFirst()[1];
      
      String[] sheets = new String[opts.getSecond().length];
      String[] excerptAllSheets = merger.getSheetNames(excerpt);
      for (int i=0; i<sheets.length; i++)
      {
         int sheetNumber = opts.getSecond()[i];
         sheets[i] = excerptAllSheets[sheetNumber];
      }
      
      File outF = new File("output.xls");
      FileOutputStream out = new FileOutputStream(outF);

      merger.merge(sheets, excerpt, full, out);
      out.close();
      
      System.out.println("Output as "  + outF);
   }
}
