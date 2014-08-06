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

import com.quanticate.opensource.spreadsheetexcerpt.excerpt.MakeReadOnlyAndExcerpt;

/**
 * CLI Tool for running the Excerpt
 */
public class SpreadsheetExcerpt extends SpreadsheetCLI
{
   public static void main(String[] args) throws Exception
   {
      Pair<File[],int[]> opts = processArgs(args, 1, "SpreadsheetExcerpt");

      File input = opts.getFirst()[0];
      int[] sheetsToKeep = opts.getSecond();

      MakeReadOnlyAndExcerpt excerpter = (MakeReadOnlyAndExcerpt)getHandler(input, true);

      File outF = new File("output.xls");
      FileOutputStream out = new FileOutputStream(outF);

      excerpter.excerpt(sheetsToKeep, input, out);
      out.close();

      System.out.println("Output as "  + outF);
   }
}
