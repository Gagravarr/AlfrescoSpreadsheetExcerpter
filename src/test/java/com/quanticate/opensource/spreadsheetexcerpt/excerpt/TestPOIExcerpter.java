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
   private POIExcerpter excerpter = new POIExcerpter();
   
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
         Row r1 = s.createRow(0);
         Cell c1 = r1.createCell(0);
         Cell c2 = r1.createCell(1);
         Cell c3 = r1.createCell(2);
         Cell c4 = r1.createCell(3);
         
         c1.setCellValue(1);
         c2.setCellValue(2);
         c3.setCellFormula("A1+B1");
         c4.setCellFormula("(A1+B1)*B1");
         // TODO String
         
         // Ensure the formulas are current
         eval.evaluateFormulaCell(c3);
         eval.evaluateFormulaCell(c4);
         
         // Run
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
      }
   }

   @Test
   public void excerptRemovesUnUsed() throws Exception
   {
      
   }

   // TODO Test excerpts
}
