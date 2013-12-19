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
import java.util.ArrayList;
import java.util.List;
import java.util.StringTokenizer;

import org.alfresco.util.Pair;

import com.quanticate.opensource.spreadsheetexcerpt.excerpt.MakeReadOnlyAndExcerpt;
import com.quanticate.opensource.spreadsheetexcerpt.excerpt.POIExcerpter;

/**
 * CLI Tool
 * 
 * TODO Spring-Enable
 */
public class SpreadsheetExcerpt
{
   public static void main(String[] args) throws Exception
   {
      Pair<File,int[]> opts = parseArgs(args);
      if (opts == null)
      {
         System.err.println("Use:");
         System.err.println("   SpreadsheetExcerpt <filename> [sheet,numbers,to,keep]");
         System.exit(1);
      }
      
      MakeReadOnlyAndExcerpt excerpter = new POIExcerpter();
      File input = opts.getFirst();
      int[] sheetsToKeep = opts.getSecond();
      
      if (sheetsToKeep.length == 0)
      {
         String[] names = excerpter.getSheetNames(input); 
         
         System.out.println("Sheets:");
         for (int i=0; i<names.length; i++)
         {
            System.out.println("  " + i + " - " + names[i]);
         }
         return;
      }
      
      File outF = new File("output.xls");
      FileOutputStream out = new FileOutputStream(outF);
      
      excerpter.excerpt(sheetsToKeep, input, out);
      out.close();
      
      System.out.println("Output as "  + outF);
   }
   
   /**
    * Process the arguments
    */
   private static Pair<File,int[]> parseArgs(String[] args)
   {
      if (args.length == 0) return null;
      if (args[0].isEmpty()) return null;
      
      File f = new File(args[0]);
      int[] sn = new int[0];
      
      if (args.length > 1 && !args[1].isEmpty())
      {
         List<Integer> sns = new ArrayList<Integer>();
         for (int i=1; i<args.length; i++)
         {
            if (args[i].length() > 0)
            {
               StringTokenizer st = new StringTokenizer(args[i], ",");
               while (st.hasMoreElements())
               {
                  sns.add(new Integer(st.nextToken().trim()));
               }
            }
         }
         
         sn = new int[sns.size()];
         for (int i=0; i<sn.length; i++)
         {
            sn[i] = sns.get(i);
         }
      }
      
      return new Pair<File,int[]>(f, sn);
   }
}
