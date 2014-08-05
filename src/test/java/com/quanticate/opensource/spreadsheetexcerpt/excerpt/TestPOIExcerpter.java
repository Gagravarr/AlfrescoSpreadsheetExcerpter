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

import static org.junit.Assert.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class TestPOIExcerpter
{
   private POIExcerpterAndMerger excerpter = new POIExcerpterAndMerger();
   
   @Test
   public void sheetListing() throws Exception
   {
      String[] names = new String[] { "a", "b", "ccc", "dddd" };
      
      for (Workbook wb : new Workbook[] { new HSSFWorkbook(), new XSSFWorkbook() })
      {
         for (String sn : names)
         {
            wb.createSheet(sn);
         }
         File tmp = File.createTempFile("test", ".xls");
         wb.write(new FileOutputStream(tmp));
         
         String[] foundNames = excerpter.getSheetNames(tmp);
         assertEquals(names.length, foundNames.length);
         for (int i=0; i<names.length; i++)
         {
            assertEquals(names[i], foundNames[i]);
         }
      }
   }
   
   @Test
   public void excerptGoesReadOnly() throws Exception
   {
      for (Workbook wb : new Workbook[] { new HSSFWorkbook(), new XSSFWorkbook() })
      {
         FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
               
         Sheet s = wb.createSheet("Test");
         
         // Numeric formulas
         Row r1 = s.createRow(0);
         Cell c1 = r1.createCell(0);
         Cell c2 = r1.createCell(1);
         Cell c3 = r1.createCell(2);
         Cell c4 = r1.createCell(3);
         
         c1.setCellValue(1);
         c2.setCellValue(2);
         c3.setCellFormula("A1+B1");
         c4.setCellFormula("(A1+B1)*B1");

         // Strings, booleans and errors
         Row r2 = s.createRow(1);
         Cell c21 = r2.createCell(0);
         Cell c22 = r2.createCell(1);
         Cell c23 = r2.createCell(2);
         Cell c24 = r2.createCell(3);
         
         c21.setCellValue("Testing");
         c22.setCellFormula("CONCATENATE(A2,A2)");
         c23.setCellFormula("FALSE()");
         c24.setCellFormula("A1/0");
         
         // Ensure the formulas are current
         eval.evaluateAll();
         
         
         // Run the excerpt
         File tmp = File.createTempFile("test", ".xls");
         wb.write(new FileOutputStream(tmp));
         
         ByteArrayOutputStream baos = new ByteArrayOutputStream();
         excerpter.excerpt(new int[] {0}, tmp, baos);
         
         
         // Check
         Workbook newwb = WorkbookFactory.create(new ByteArrayInputStream(baos.toByteArray()));
         assertEquals(1, newwb.getNumberOfSheets());
         
         s = newwb.getSheetAt(0);
         r1 = s.getRow(0);
         assertEquals(Cell.CELL_TYPE_NUMERIC, r1.getCell(0).getCellType());
         assertEquals(Cell.CELL_TYPE_NUMERIC, r1.getCell(1).getCellType());
         assertEquals(Cell.CELL_TYPE_NUMERIC, r1.getCell(2).getCellType());
         assertEquals(Cell.CELL_TYPE_NUMERIC, r1.getCell(3).getCellType());

         assertEquals(1.0, s.getRow(0).getCell(0).getNumericCellValue(), 0.001);
         assertEquals(2.0, s.getRow(0).getCell(1).getNumericCellValue(), 0.001);
         assertEquals(3.0, s.getRow(0).getCell(2).getNumericCellValue(), 0.001);
         assertEquals(6.0, s.getRow(0).getCell(3).getNumericCellValue(), 0.001);
         
         r2 = s.getRow(1);
         assertEquals(Cell.CELL_TYPE_STRING,  r2.getCell(0).getCellType());
         assertEquals(Cell.CELL_TYPE_STRING,  r2.getCell(1).getCellType());
         assertEquals(Cell.CELL_TYPE_BOOLEAN, r2.getCell(2).getCellType());
         assertEquals(Cell.CELL_TYPE_BLANK,   r2.getCell(3).getCellType());

         assertEquals("Testing",        s.getRow(1).getCell(0).getStringCellValue());
         assertEquals("TestingTesting", s.getRow(1).getCell(1).getStringCellValue());
         assertEquals(false, s.getRow(1).getCell(2).getBooleanCellValue());
      }
   }

   @Test
   public void excerptRemovesUnUsed() throws Exception
   {
      String[] names = new String[] { "a", "b", "ccc", "dddd", "e", "f", "gg" };
      
      for (Workbook wb : new Workbook[] { new HSSFWorkbook(), new XSSFWorkbook() })
      {
         // Create some dummy content
         for (String sn : names)
         {
            Sheet s = wb.createSheet(sn);
            s.createRow(0).createCell(0).setCellValue(sn);
         }

         
         // Excerpt by index
         File tmp = File.createTempFile("test", ".xls");
         wb.write(new FileOutputStream(tmp));
         
         ByteArrayOutputStream baos = new ByteArrayOutputStream();
         int[] excI = new int[] {0, 1, 2, 4, 5};
         excerpter.excerpt(excI, tmp, baos);

         // Check
         Workbook newwb = WorkbookFactory.create(new ByteArrayInputStream(baos.toByteArray()));
         assertEquals(5, newwb.getNumberOfSheets());
         assertEquals(names[excI[0]], newwb.getSheetName(0));
         assertEquals(names[excI[1]], newwb.getSheetName(1));
         assertEquals(names[excI[2]], newwb.getSheetName(2));
         assertEquals(names[excI[3]], newwb.getSheetName(3));
         assertEquals(names[excI[4]], newwb.getSheetName(4));
         
         assertEquals(names[excI[0]], newwb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
         assertEquals(names[excI[1]], newwb.getSheetAt(1).getRow(0).getCell(0).getStringCellValue());
         assertEquals(names[excI[2]], newwb.getSheetAt(2).getRow(0).getCell(0).getStringCellValue());
         assertEquals(names[excI[3]], newwb.getSheetAt(3).getRow(0).getCell(0).getStringCellValue());
         assertEquals(names[excI[4]], newwb.getSheetAt(4).getRow(0).getCell(0).getStringCellValue());
         
         
         // Excerpt by name
         String[] excN = new String[] { "b", "ccc", "f", "gg" };
         baos = new ByteArrayOutputStream();
         excerpter.excerpt(excN, tmp, baos);
         
         newwb = WorkbookFactory.create(new ByteArrayInputStream(baos.toByteArray()));
         assertEquals(4, newwb.getNumberOfSheets());
         assertEquals(excN[0], newwb.getSheetName(0));
         assertEquals(excN[1], newwb.getSheetName(1));
         assertEquals(excN[2], newwb.getSheetName(2));
         assertEquals(excN[3], newwb.getSheetName(3));
         
         assertEquals(excN[0], newwb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
         assertEquals(excN[1], newwb.getSheetAt(1).getRow(0).getCell(0).getStringCellValue());
         assertEquals(excN[2], newwb.getSheetAt(2).getRow(0).getCell(0).getStringCellValue());
         assertEquals(excN[3], newwb.getSheetAt(3).getRow(0).getCell(0).getStringCellValue());

         
         // Can't excerpt by invalid index
         try
         {
            excerpter.excerpt(new int[] {0, 10}, tmp, null);
            fail();
         }
         catch (IllegalArgumentException e) {}
         
         // Can't excerpt by invalid name
         try
         {
            excerpter.excerpt(new String[] {"a", "invalid"}, tmp, null);
            fail();
         }
         catch (IllegalArgumentException e) {}
      }
   }
}
