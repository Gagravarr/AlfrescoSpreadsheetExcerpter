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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.alfresco.repo.content.MimetypeMap;
import org.alfresco.service.cmr.repository.ContentReader;
import org.alfresco.service.cmr.repository.ContentWriter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.quanticate.opensource.spreadsheetexcerpt.merge.MergeChangesFromExcerpt;

public class POIExcerpterAndMerger implements MakeReadOnlyAndExcerpt, MergeChangesFromExcerpt
{
   private static final Set<String> SUPPORTED_MIMETYPES = 
         Collections.unmodifiableSet(new HashSet<String>(Arrays.asList(new String[] { 
               MimetypeMap.MIMETYPE_EXCEL, 
               MimetypeMap.MIMETYPE_OPENDOCUMENT_SPREADSHEET,
               MimetypeMap.MIMETYPE_OPENDOCUMENT_SPREADSHEET_TEMPLATE
         })));

   @Override
   public Set<String> getSupportedMimeTypes()
   {
      return SUPPORTED_MIMETYPES;
   }

   private Workbook open(File f) throws IOException
   {
      try
      {
         return WorkbookFactory.create(f);
      }
      catch (InvalidFormatException e)
      {
         throw new IOException("File broken", e);
      }
   }
   private Workbook open(ContentReader reader) throws IOException
   {
      // If we can use a FileChannel, do
      if (reader.getMimetype().equals(MimetypeMap.MIMETYPE_EXCEL))
      {
         NPOIFSFileSystem fs = new NPOIFSFileSystem(reader.getFileChannel());
         return WorkbookFactory.create(fs);
      }

      // Otherwise, ContentReader doesn't offer a File
      // So, we have to go via the InputStream
      try
      {
         return WorkbookFactory.create(reader.getContentInputStream());
      }
      catch (InvalidFormatException e)
      {
         throw new IOException("File broken", e);
      }
   }
   private OutputStream open(ContentWriter dest, ContentReader source)
   {
      // Copy over the mimetype
      dest.setMimetype(source.getMimetype());

      // Now get the output
      return dest.getContentOutputStream();
   }


   // =================================================================== 


   @Override
   public String[] getSheetNames(File f) throws IOException
   {
      Workbook wb = open(f);

      String[] names = new String[wb.getNumberOfSheets()];
      for (int i=0; i<names.length; i++)
      {
         names[i] = wb.getSheetName(i);
      }
      return names;
   }

   @Override
   public void excerpt(String[] sheetsToKeep, File input, OutputStream output) throws IOException
   {
      Workbook wb = open(input);
      excerpt(sheetsToKeep, wb, output);
   }

   @Override
   public void excerpt(String[] sheetsToKeep, ContentReader input, ContentWriter output) throws IOException
   {
      Workbook wb = open(input);

      OutputStream stream = open(output, input);
      excerpt(sheetsToKeep, wb, stream);
      stream.close();
   }

   protected void excerpt(String[] sheetsToKeep, Workbook wb, OutputStream output) throws IOException
   {
      List<Sheet> keep = new ArrayList<Sheet>(sheetsToKeep.length);
      for (String sn : sheetsToKeep)
      {
         Sheet s = wb.getSheet(sn);
         if (s == null)
            throw new IllegalArgumentException("Sheet not found with name '" + sn + "'");
            
         keep.add(s);
      }
      
      excerpt(wb, keep, output);
   }

   @Override
   public void excerpt(int[] sheetsToKeep, File input, OutputStream output) throws IOException
   {
      Workbook wb = open(input);
      excerpt(sheetsToKeep, wb, output);
   }
      
   @Override
   public void excerpt(int[] sheetsToKeep, ContentReader input, ContentWriter output) throws IOException
   {
      Workbook wb = open(input);

      OutputStream stream = open(output, input);
      excerpt(sheetsToKeep, wb, stream);
      stream.close();
   }


   // =================================================================== 


   @Override
   public void merge(String[] sheetsToMerge, File excerptInput, File fullInput, OutputStream output) throws IOException
   {
      Workbook excerptWB = open(excerptInput);
      Workbook fullWB = open(fullInput);

      merge(excerptWB, fullWB, sheetsToMerge, output);
   }

   @Override
   public void excerpt(String[] sheetsToMerge, ContentReader excerptInput, ContentReader fullInput, ContentWriter output) throws IOException
   {
      Workbook excerptWB = open(excerptInput);
      Workbook fullWB = open(fullInput);

      OutputStream stream = open(output, fullInput);
      merge(excerptWB, fullWB, sheetsToMerge, stream);
      stream.close();
   }


   // =================================================================== 


   protected void excerpt(int[] sheetsToKeep, Workbook wb, OutputStream output) throws IOException
   {
      List<Sheet> keep = new ArrayList<Sheet>(sheetsToKeep.length);
      for (int sn : sheetsToKeep)
      {
         keep.add( wb.getSheetAt(sn) );
      }

      excerpt(wb, keep, output);
   }

   private void excerpt(Workbook wb, List<Sheet> sheetsToKeep, OutputStream output) throws IOException
   {
      // Make the requested sheets be read only
      Set<String> keepNames = new HashSet<String>();
      for (Sheet s : sheetsToKeep)
      {
         keepNames.add(s.getSheetName());
         for (Row r : s)
         {
            for (Cell c : r)
            {
               if (c.getCellType() == Cell.CELL_TYPE_FORMULA)
               {
                  switch (c.getCachedFormulaResultType())
                  {
                     case Cell.CELL_TYPE_NUMERIC:
                        double vd = c.getNumericCellValue();
                        c.setCellType(Cell.CELL_TYPE_NUMERIC);
                        c.setCellValue(vd);
                        break;
                     case Cell.CELL_TYPE_STRING:
                        RichTextString vs = c.getRichStringCellValue();
                        c.setCellType(Cell.CELL_TYPE_STRING);
                        c.setCellValue(vs);
                        break;
                     case Cell.CELL_TYPE_BOOLEAN:
                        boolean vb = c.getBooleanCellValue();
                        c.setCellType(Cell.CELL_TYPE_BOOLEAN);
                        c.setCellValue(vb);
                        break;
                     case Cell.CELL_TYPE_ERROR:
                        c.setCellType(Cell.CELL_TYPE_BLANK);
                        break;
                  }
               }
            }
         }
      }
      
      // Remove all the other sheets
      // Note - work backwards! Avoids order changing under us
      for (int i=wb.getNumberOfSheets()-1; i>=0; i--)
      {
         String name = wb.getSheetName(i);
         if (! keepNames.contains(name))
         {
            wb.removeSheetAt(i);
         }
      }
      
      // Save
      wb.write(output);
   }


   // =================================================================== 

   private void merge(Workbook excerptWB, Workbook fullWB, String[] sheetsToMerge, OutputStream output) throws IOException
   {
      // TODO Refactor / re-order String[] to List<> logic
      // TODO Implement
   }

   private void merge(Workbook excerptWB, Workbook fullWB, List<Sheet> sheetsToMerge, OutputStream output) throws IOException
   {
      // TODO Implement
   }
}
