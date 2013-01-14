package com.quanticate.opensource.spreadsheetexcerpt.excerpt;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class POIExcerpter implements MakeReadOnlyAndExcerpt
{
   @Override
   public Set<String> getSupportedMimeTypes()
   {
      return null; // TODO From MimeypesMap
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
      
      List<Sheet> keep = new ArrayList<Sheet>(sheetsToKeep.length);
      for (String sn : sheetsToKeep)
      {
         keep.add( wb.getSheet(sn) );
      }
      
      excerpt(wb, keep, output);
   }

   @Override
   public void excerpt(int[] sheetsToKeep, File input, OutputStream output) throws IOException
   {
      Workbook wb = open(input);
      
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
                        c.setCellValue( c.getNumericCellValue() );
                     case Cell.CELL_TYPE_STRING:
                        c.setCellValue( c.getRichStringCellValue() );
                     case Cell.CELL_TYPE_BOOLEAN:
                        c.setCellValue( c.getBooleanCellValue() );
                     case Cell.CELL_TYPE_ERROR:
                        c.setCellType(Cell.CELL_TYPE_BLANK);
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
}
