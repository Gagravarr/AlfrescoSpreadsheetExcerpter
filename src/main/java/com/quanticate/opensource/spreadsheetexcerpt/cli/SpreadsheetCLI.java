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
import java.util.ArrayList;
import java.util.List;
import java.util.StringTokenizer;

import org.alfresco.util.Pair;

/**
 * Parent for the CLI tools
 */
public abstract class SpreadsheetCLI
{
   public static Pair<File[], int[]> processArgs(String[] args, int expectedFiles, String name)
   {
      Pair<File[], int[]> opts = parseArgs(args, expectedFiles);

      if (opts == null || opts.getFirst().length == 0)
      {
         StringBuffer sb = new StringBuffer();
         sb.append("   ");
         sb.append(name);
         for (int i=0; i<expectedFiles; i++) {
            sb.append(" <filename>");
         }
         sb.append(" [sheet,numbers,to,use]");

         System.err.println("Use:");
         System.err.println(sb.toString());
         System.exit(1);
      }

      if (opts.getSecond().length == 0)
      {
         // TODO Implement listing
//         String[] names = excerpter.getSheetNames(input); 
         String[] names = new String[0]; 

         System.out.println("Sheets:");
         for (int i=0; i<names.length; i++)
         {
            System.out.println("  " + i + " - " + names[i]);
         }
         System.exit(0);
      }

      return opts;
   }

   /**
    * Process the arguments
    */
   private static Pair<File[],int[]> parseArgs(String[] args, int expectedFiles)
   {
      if (args.length < expectedFiles) return null;

      File[] files = new File[expectedFiles];
      for (int i=0; i<expectedFiles; i++) 
      {
         if (args[i].isEmpty()) return null;
         
         files[i] = new File(args[i]);
         if (! files[i].exists())
         {
            System.err.println("File not found - " + files[i]);
            System.exit(2);
         }
      }
      
      int[] sn = new int[0];
      
      if (args.length > expectedFiles && !args[expectedFiles].isEmpty())
      {
         List<Integer> sns = new ArrayList<Integer>();
         for (int i=expectedFiles; i<args.length; i++)
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
      
      return new Pair<File[],int[]>(files, sn);
   }
}
